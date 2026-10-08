import contextlib
import datetime
import io
import json
import os
import shutil
import subprocess
import sys

from nxdoc import NxDoc


# find kubectl: in the PATH, or next to this script (kubectl.exe on Windows)
def find_kubectl():
    kubectl_path = shutil.which("kubectl")
    if kubectl_path is not None:
        return kubectl_path

    script_folder = os.path.dirname(os.path.abspath(__file__))
    for file_name in ["kubectl.exe", "kubectl"]:
        local_kubectl = os.path.join(script_folder, file_name)
        if os.path.isfile(local_kubectl):
            return local_kubectl

    print("kubectl not found.")
    print("Install kubectl, or copy kubectl.exe into: " + script_folder)
    sys.exit(1)


# run a kubectl command against the management cluster and return the parsed json
def run_kubectl(kubeconfig_path, arguments):
    command = [find_kubectl(), "--kubeconfig", kubeconfig_path] + arguments + ["-o", "json"]
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        print("kubectl failed: " + " ".join(command))
        print(result.stderr)
        sys.exit(1)
    return json.loads(result.stdout)


# save the collected data so we can re-test offline
def save_inventory(inventory, file_name):
    with open(file_name, "w") as inventory_file:
        json.dump(inventory, inventory_file, indent=2)


# find the clusterConfig variable of a cluster
def get_cluster_config(cluster):
    variables = cluster["spec"]["topology"].get("variables", [])
    for variable in variables:
        if variable["name"] == "clusterConfig":
            return variable["value"]
    return {}


# find the workerConfig override of a worker pool
def get_worker_config(worker_pool):
    overrides = worker_pool.get("variables", {}).get("overrides", [])
    for override in overrides:
        if override["name"] == "workerConfig":
            return override["value"]
    return {}


# read the vm sizing from a nutanix machineDetails block
def read_machine_sizing(machine_details):
    subnet_names = []
    for subnet in machine_details.get("subnets", []):
        subnet_names.append(subnet.get("name", ""))

    sizing = {}
    sizing["pe_cluster"] = machine_details.get("cluster", {}).get("name", "")
    sizing["image_name"] = machine_details.get("image", {}).get("name", "")
    sizing["memory_size"] = machine_details.get("memorySize", "")
    sizing["disk_size"] = machine_details.get("systemDiskSize", "")
    sizing["vcpu_sockets"] = machine_details.get("vcpuSockets", "")
    sizing["vcpus_per_socket"] = machine_details.get("vcpusPerSocket", "")
    sizing["subnets"] = ", ".join(subnet_names)
    return sizing


# get the replicas of a worker pool, or min / max when the autoscaler is used
def get_pool_replicas(worker_pool):
    if "replicas" in worker_pool:
        return str(worker_pool["replicas"])

    annotations = worker_pool.get("metadata", {}).get("annotations", {})
    min_size = annotations.get("cluster.x-k8s.io/cluster-api-autoscaler-node-group-min-size", "")
    max_size = annotations.get("cluster.x-k8s.io/cluster-api-autoscaler-node-group-max-size", "")
    if min_size == "" and max_size == "":
        return ""
    return "min " + min_size + " / max " + max_size


# read the control plane and worker pool sizing of a cluster
def read_cluster_sizing(cluster):
    topology = cluster["spec"]["topology"]
    cluster_config = get_cluster_config(cluster)

    # control plane sizing
    control_plane_details = cluster_config.get("controlPlane", {}).get("nutanix", {}).get("machineDetails", {})
    control_plane = read_machine_sizing(control_plane_details)
    control_plane["replicas"] = str(topology.get("controlPlane", {}).get("replicas", ""))

    # worker pools sizing
    worker_pools = []
    machine_deployments = topology.get("workers", {}).get("machineDeployments", [])
    for worker_pool in machine_deployments:
        worker_config = get_worker_config(worker_pool)
        worker_details = worker_config.get("nutanix", {}).get("machineDetails", {})
        pool_sizing = read_machine_sizing(worker_details)
        pool_sizing["pool_name"] = worker_pool["name"]
        pool_sizing["replicas"] = get_pool_replicas(worker_pool)
        worker_pools.append(pool_sizing)

    cluster_sizing = {}
    cluster_sizing["control_plane"] = control_plane
    cluster_sizing["worker_pools"] = worker_pools
    return cluster_sizing


# get the internal ip address of a machine
def get_machine_ip(machine):
    addresses = machine.get("status", {}).get("addresses", [])
    for address in addresses:
        if address["type"] == "InternalIP":
            return address["address"]
    return ""


# read the nodes (machines) of one cluster
def read_cluster_nodes(cluster_name, machines_data):
    nodes = []
    for machine in machines_data["items"]:
        labels = machine["metadata"].get("labels", {})
        if labels.get("cluster.x-k8s.io/cluster-name") != cluster_name:
            continue

        # use the kubernetes node name, or the machine name if the node is not ready yet
        node_name = machine.get("status", {}).get("nodeRef", {}).get("name", "")
        if node_name == "":
            node_name = machine["metadata"]["name"]

        node = {}
        node["node_name"] = node_name
        node["ip_address"] = get_machine_ip(machine)
        if "cluster.x-k8s.io/control-plane" in labels:
            node["role"] = "control-plane"
            node["pool_name"] = ""
        else:
            node["role"] = "worker"
            node["pool_name"] = labels.get("topology.cluster.x-k8s.io/deployment-name", "")
        nodes.append(node)
    return nodes


# ask a question, use the default answer when the user just presses Enter
def ask(question, default_answer=""):
    if default_answer != "":
        question = question + " [" + default_answer + "]"
    answer = input(question + ": ").strip()
    if answer == "":
        answer = default_answer
    return answer


# ask the user everything kubectl cannot give
def ask_document_details():
    details = {}
    details["customer_name"] = ask("Customer name")
    details["project_name"] = ask("Project name")
    details["author"] = ask("Author")
    details["document_version"] = ask("Document version", "1.0")
    details["kubeconfig_path"] = ask("Management cluster kubeconfig path", "nkp-mgmt.conf")
    return details


# chapter 1: document control
def add_document_control(doc, details, today):
    doc.h1("Document Control")

    # document information
    doc.h2("Document Information")
    information_rows = []
    information_rows.append(["Customer", details["customer_name"]])
    information_rows.append(["Project", details["project_name"]])
    information_rows.append(["Document Title", "NKP As-built Guide"])
    information_rows.append(["Document Version", details["document_version"]])
    information_rows.append(["Author", details["author"]])
    information_rows.append(["Date", today])
    doc.table(["Item", "Value"], information_rows, widths=[1, 2],
              caption="Document information", bold_first_col=True)

    # revision history
    doc.h2("Revision History")
    revision_rows = []
    revision_rows.append([details["document_version"], today, details["author"], "Initial release"])
    doc.table(["Version", "Date", "Author", "Description"], revision_rows, widths=[1, 1.5, 2, 3],
              caption="Revision history")


# check if a cluster is the management (kommander host) cluster
def is_management_cluster(cluster):
    labels = cluster["metadata"].get("labels", {})
    return labels.get("kommander.d2iq.io/host") == "true"


# count the control plane and worker nodes of a cluster
def count_nodes(nodes):
    control_plane_count = 0
    worker_count = 0
    for node in nodes:
        if node["role"] == "control-plane":
            control_plane_count = control_plane_count + 1
        else:
            worker_count = worker_count + 1
    return control_plane_count, worker_count


# chapter 2: architecture summary
def add_architecture_summary(doc, clusters_data, machines_data):
    doc.h1("Architecture Summary")

    # find the management cluster name and the kubernetes versions in use
    management_cluster_name = ""
    workload_cluster_count = 0
    kubernetes_versions = []
    for cluster in clusters_data["items"]:
        if is_management_cluster(cluster):
            management_cluster_name = cluster["metadata"]["name"]
        else:
            workload_cluster_count = workload_cluster_count + 1
        kubernetes_version = cluster["spec"]["topology"].get("version", "")
        if kubernetes_version not in kubernetes_versions:
            kubernetes_versions.append(kubernetes_version)

    # short text
    doc.h2("Overview")
    doc.p("The Nutanix Kubernetes Platform (NKP) environment is made of one management cluster ("
          + management_cluster_name + ") and " + str(workload_cluster_count)
          + " workload clusters, deployed on Nutanix infrastructure.")
    doc.p("Kubernetes version in use: " + ", ".join(kubernetes_versions) + ".")

    # architecture diagram placeholder
    doc.h2("Architecture Diagram")
    doc.instructions("CONSULTANT - Insert the architecture diagram here")
    doc.placeholder_figure()

    # table of all clusters
    doc.h2("Clusters")
    cluster_rows = []
    for cluster in clusters_data["items"]:
        cluster_name = cluster["metadata"]["name"]
        if is_management_cluster(cluster):
            cluster_role = "Management"
        else:
            cluster_role = "Workload"
        nodes = read_cluster_nodes(cluster_name, machines_data)
        control_plane_count, worker_count = count_nodes(nodes)
        cluster_rows.append([
            cluster_name,
            cluster_role,
            cluster["metadata"]["namespace"],
            cluster["spec"]["topology"].get("version", ""),
            cluster["spec"].get("controlPlaneEndpoint", {}).get("host", ""),
            str(control_plane_count),
            str(worker_count),
        ])
    doc.table(["Cluster", "Role", "Namespace", "Kubernetes", "API Endpoint", "CP Nodes", "Workers"],
              cluster_rows, widths=[2.2, 1.4, 3, 1.3, 1.6, 0.9, 1], caption="NKP clusters")


# find the management cluster in the cluster list
def get_management_cluster(clusters_data):
    for cluster in clusters_data["items"]:
        if is_management_cluster(cluster):
            return cluster
    return None


# build the dashboard url from the traefik load balancer ip and the dashboard ingress path
def get_dashboard_url(traefik_service, dashboard_ingress):
    load_balancer = traefik_service.get("status", {}).get("loadBalancer", {}).get("ingress", [])
    if len(load_balancer) == 0:
        return ""
    dashboard_address = load_balancer[0].get("hostname", "")
    if dashboard_address == "":
        dashboard_address = load_balancer[0].get("ip", "")

    dashboard_path = ""
    for rule in dashboard_ingress["spec"].get("rules", []):
        for ingress_path in rule.get("http", {}).get("paths", []):
            dashboard_path = ingress_path.get("path", "")
    return "https://" + dashboard_address + dashboard_path


# chapter 3: nkp platform
def add_nkp_platform(doc, platform_data, clusters_data):
    doc.h1("NKP Platform")

    # nkp version and air-gapped mode
    nkp_version = ""
    air_gapped = ""
    for kommander_core in platform_data["kommander_cores"]["items"]:
        nkp_version = kommander_core["spec"].get("version", "")
        if kommander_core["spec"].get("airgapped", {}).get("enabled", False):
            air_gapped = "Yes"
        else:
            air_gapped = "No"

    # license level
    license_text = ""
    for nkp_license in platform_data["licenses"]["items"]:
        license_status = nkp_license.get("status", {})
        license_text = license_status.get("productName", "") + " " + license_status.get("dkpLevel", "")
        if license_status.get("valid", False):
            license_text = license_text + " (valid)"
        else:
            license_text = license_text + " (not valid)"

    # management cluster and its registry mirror
    management_cluster_name = ""
    registry_mirror = ""
    management_cluster = get_management_cluster(clusters_data)
    if management_cluster is not None:
        management_cluster_name = management_cluster["metadata"]["name"]
        cluster_config = get_cluster_config(management_cluster)
        registry_mirror = cluster_config.get("globalImageRegistryMirror", {}).get("url", "")

    dashboard_url = get_dashboard_url(platform_data["traefik_service"], platform_data["dashboard_ingress"])

    doc.h2("Platform Details")
    platform_rows = []
    platform_rows.append(["NKP Version", nkp_version])
    platform_rows.append(["License", license_text])
    platform_rows.append(["Management Cluster", management_cluster_name])
    platform_rows.append(["Air-gapped", air_gapped])
    platform_rows.append(["Image Registry Mirror", registry_mirror])
    platform_rows.append(["Dashboard URL", dashboard_url])
    doc.table(["Item", "Value"], platform_rows, widths=[1, 2],
              caption="NKP platform details", bold_first_col=True)


# add a value to a list only if it is not empty and not already there
def add_unique(values, value):
    if value != "" and value not in values:
        values.append(value)


# get the nutanix storage container used by the csi storage classes
def get_storage_container(cluster_config):
    nutanix_csi = cluster_config.get("addons", {}).get("csi", {}).get("providers", {}).get("nutanix", {})
    storage_class_configs = nutanix_csi.get("storageClassConfigs", {})
    storage_containers = []
    for storage_class_name in storage_class_configs:
        parameters = storage_class_configs[storage_class_name].get("parameters", {})
        add_unique(storage_containers, parameters.get("storageContainer", ""))
    return ", ".join(storage_containers)


# chapter 4: nutanix infrastructure
def add_nutanix_infrastructure(doc, clusters_data):
    doc.h1("Nutanix Infrastructure")

    # collect the values used by all clusters
    prism_central_urls = []
    pe_clusters = []
    subnets = []
    storage_containers = []
    os_images = []
    ntp_servers = []
    for cluster in clusters_data["items"]:
        cluster_config = get_cluster_config(cluster)
        prism_central = cluster_config.get("nutanix", {}).get("prismCentralEndpoint", {})
        add_unique(prism_central_urls, prism_central.get("url", ""))
        add_unique(storage_containers, get_storage_container(cluster_config))
        for ntp_server in cluster_config.get("ntp", {}).get("servers", []):
            add_unique(ntp_servers, ntp_server)

        # pe cluster, subnets and image of the control plane and worker pools
        cluster_sizing = read_cluster_sizing(cluster)
        all_sizing = [cluster_sizing["control_plane"]] + cluster_sizing["worker_pools"]
        for sizing in all_sizing:
            add_unique(pe_clusters, sizing["pe_cluster"])
            add_unique(os_images, sizing["image_name"])
            for subnet_name in sizing["subnets"].split(", "):
                add_unique(subnets, subnet_name)

    doc.h2("Infrastructure Details")
    infrastructure_rows = []
    infrastructure_rows.append(["Prism Central", ", ".join(prism_central_urls)])
    infrastructure_rows.append(["Prism Element Clusters", ", ".join(pe_clusters)])
    infrastructure_rows.append(["Subnets", ", ".join(subnets)])
    infrastructure_rows.append(["Storage Containers", ", ".join(storage_containers)])
    infrastructure_rows.append(["OS Images", ", ".join(os_images)])
    infrastructure_rows.append(["NTP Servers", ", ".join(ntp_servers)])
    doc.table(["Item", "Value"], infrastructure_rows, widths=[1, 2],
              caption="Nutanix infrastructure details", bold_first_col=True)


# make the sizing rows of a control plane or worker pool table
def make_sizing_rows(sizing):
    sizing_rows = []
    sizing_rows.append(["Replicas", sizing["replicas"]])
    sizing_rows.append(["Prism Element Cluster", sizing["pe_cluster"]])
    sizing_rows.append(["OS Image", sizing["image_name"]])
    sizing_rows.append(["vCPU (Sockets x Cores)", str(sizing["vcpu_sockets"]) + " x " + str(sizing["vcpus_per_socket"])])
    sizing_rows.append(["Memory", sizing["memory_size"]])
    sizing_rows.append(["System Disk", sizing["disk_size"]])
    sizing_rows.append(["Subnets", sizing["subnets"]])
    return sizing_rows


# chapter per cluster: general, networking and storage, control plane, worker pools, nodes
def add_cluster_chapter(doc, cluster, machines_data):
    cluster_name = cluster["metadata"]["name"]
    topology = cluster["spec"]["topology"]
    cluster_config = get_cluster_config(cluster)
    addons = cluster_config.get("addons", {})
    doc.h1("Cluster " + cluster_name)

    # general
    if is_management_cluster(cluster):
        cluster_role = "Management"
    else:
        cluster_role = "Workload"
    control_plane_endpoint = cluster["spec"].get("controlPlaneEndpoint", {})
    endpoint_text = ""
    if control_plane_endpoint.get("host", "") != "":
        endpoint_text = control_plane_endpoint["host"] + ":" + str(control_plane_endpoint.get("port", ""))

    doc.h2("General")
    general_rows = []
    general_rows.append(["Cluster Name", cluster_name])
    general_rows.append(["Namespace", cluster["metadata"]["namespace"]])
    general_rows.append(["Role", cluster_role])
    general_rows.append(["Provider", cluster["metadata"].get("labels", {}).get("cluster.x-k8s.io/provider", "")])
    general_rows.append(["Kubernetes Version", topology.get("version", "")])
    general_rows.append(["Cluster Class", topology.get("classRef", {}).get("name", "")])
    general_rows.append(["Control Plane Endpoint", endpoint_text])
    doc.table(["Item", "Value"], general_rows, widths=[1, 2],
              caption=cluster_name + " general details", bold_first_col=True)

    # networking and storage
    cluster_network = cluster["spec"].get("clusterNetwork", {})
    pod_cidrs = cluster_network.get("pods", {}).get("cidrBlocks", [])
    service_cidrs = cluster_network.get("services", {}).get("cidrBlocks", [])

    service_load_balancer = addons.get("serviceLoadBalancer", {})
    address_ranges = []
    for address_range in service_load_balancer.get("configuration", {}).get("addressRanges", []):
        address_ranges.append(address_range.get("start", "") + " to " + address_range.get("end", ""))

    image_registries = []
    for image_registry in cluster_config.get("imageRegistries", []):
        add_unique(image_registries, image_registry.get("url", ""))

    doc.h2("Networking and Storage")
    network_rows = []
    network_rows.append(["CNI Provider", addons.get("cni", {}).get("provider", "")])
    network_rows.append(["Pod CIDR", ", ".join(pod_cidrs)])
    network_rows.append(["Service CIDR", ", ".join(service_cidrs)])
    network_rows.append(["Service Load Balancer", service_load_balancer.get("provider", "")])
    network_rows.append(["Load Balancer Address Range", ", ".join(address_ranges)])
    network_rows.append(["Storage Container", get_storage_container(cluster_config)])
    network_rows.append(["Global Image Registry Mirror", cluster_config.get("globalImageRegistryMirror", {}).get("url", "")])
    network_rows.append(["Image Registries", ", ".join(image_registries)])
    doc.table(["Item", "Value"], network_rows, widths=[1, 2],
              caption=cluster_name + " networking and storage", bold_first_col=True)

    # control plane sizing
    cluster_sizing = read_cluster_sizing(cluster)
    doc.h2("Control Plane")
    doc.table(["Item", "Value"], make_sizing_rows(cluster_sizing["control_plane"]), widths=[1, 2],
              caption=cluster_name + " control plane", bold_first_col=True)

    # worker pools sizing, one table per pool
    doc.h2("Worker Node Pools")
    for pool_sizing in cluster_sizing["worker_pools"]:
        doc.table(["Item", "Value"], make_sizing_rows(pool_sizing), widths=[1, 2],
                  caption=cluster_name + " worker pool " + pool_sizing["pool_name"], bold_first_col=True)

    # nodes: control plane first, then workers
    nodes = read_cluster_nodes(cluster_name, machines_data)
    control_plane_rows = []
    worker_rows = []
    for node in nodes:
        if node["role"] == "control-plane":
            control_plane_rows.append([node["node_name"], "Control Plane", "", node["ip_address"]])
        else:
            worker_rows.append([node["node_name"], "Worker", node["pool_name"], node["ip_address"]])
    control_plane_rows.sort()
    worker_rows.sort()

    doc.h2("Nodes")
    doc.table(["Node Name", "Role", "Pool", "IP Address"], control_plane_rows + worker_rows,
              widths=[3.2, 1.3, 0.8, 1.4], caption=cluster_name + " nodes")


# main flow
details = ask_document_details()
kubeconfig_path = details["kubeconfig_path"]

# collect the data from the management cluster
clusters_data = run_kubectl(kubeconfig_path, ["get", "clusters.cluster.x-k8s.io", "-A"])
machines_data = run_kubectl(kubeconfig_path, ["get", "machines.cluster.x-k8s.io", "-A"])

platform_data = {}
platform_data["kommander_cores"] = run_kubectl(kubeconfig_path, ["get", "kommandercores"])
platform_data["licenses"] = run_kubectl(kubeconfig_path, ["get", "licenses", "-n", "kommander"])
platform_data["traefik_service"] = run_kubectl(kubeconfig_path, ["get", "service", "kommander-traefik", "-n", "kommander"])
platform_data["dashboard_ingress"] = run_kubectl(kubeconfig_path, ["get", "ingress", "kommander-kommander-ui", "-n", "kommander"])

inventory = {}
inventory["clusters"] = clusters_data
inventory["machines"] = machines_data
inventory["platform"] = platform_data
save_inventory(inventory, "nkp-inventory.json")

# build the document
today = datetime.date.today().strftime("%d %B %Y")
doc = NxDoc("asbuilt", title="NKP As-built Guide", subtitle=details["customer_name"],
            author=details["author"])
add_document_control(doc, details, today)
add_architecture_summary(doc, clusters_data, machines_data)
add_nkp_platform(doc, platform_data, clusters_data)
add_nutanix_infrastructure(doc, clusters_data)

# one chapter per cluster, management cluster first
for cluster in clusters_data["items"]:
    if is_management_cluster(cluster):
        add_cluster_chapter(doc, cluster, machines_data)
for cluster in clusters_data["items"]:
    if not is_management_cluster(cluster):
        add_cluster_chapter(doc, cluster, machines_data)

document_name = details["customer_name"] + " - NKP As-built Guide.docx"
# no "update fields" prompt in Word, update the contents manually
# hide the technical messages printed while saving
with contextlib.redirect_stdout(io.StringIO()):
    doc.save(document_name, update_fields_on_open=False)
