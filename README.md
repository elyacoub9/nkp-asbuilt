# nkp-asbuilt

generate an **As-built Guide** for Nutanix Kubernetes Platform (NKP).

the script connects to the NKP management cluster with kubectl, collects the clusters info, then builds the document.

the document contains:
- **Document Control**
- **Architecture Summary**
- **NKP Platform**
- **Nutanix Infrastructure**
- **Management Cluster** (general, networking and storage, control plane, worker pools, nodes)
- **Workload Clusters** (same as management cluster)



## requirements

- python 3.9 or newer
- kubectl
- kubeconfig of the management cluster, **read-only** is enough 



## install

1. clone this repository:

   ```bash
   git clone https://github.com/elyacoub9/nkp-asbuilt.git
   ```

2. change to the project directory:

   ```bash
   cd nkp-asbuilt
   ```

3. install the packages:

   linux / macOS:

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

   windows (PowerShell):

   ```powershell
   python -m venv .venv
   .venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   ```

   if PowerShell blocks the activate script, run this once, then activate again:

   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
   ```

4. copy the kubeconfig of the management cluster into the project directory.



## run

run your code

linux / macOS:

```bash
python3 nkp_asbuilt.py
```

windows (in the PowerShell window where `(.venv)` is shown):

```powershell
python nkp_asbuilt.py
```

the script will ask you some questions

open the .docx in Word after the script finishes
