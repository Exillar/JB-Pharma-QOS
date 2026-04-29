# QOS

Generates the Quality Overall Summary (QOS) DOCX from dossier PDFs.

## Run

1) Copy `config.yml.example` to `config.yaml` and update the paths.
	Example:

	```yaml
	template_docx: 'D:\JB Pharma internal\JB Pharma\Tagging- mapping logic\Tagging- mapping logic\Quality Overall Summary.docx'
	dossier_root: 'D:\JB Pharma internal\JB Pharma\Cardiolek'
	filled_reference_docx: 'D:\JB Pharma internal\JB Pharma\Cardiolek\Quality Overall Summary_QOS.docx'
	output_docx: 'D:\JB Pharma internal\JB Pharma\Cardiolek\Output\Quality Overall Summary_QOS_2.3.S.1_Auto.docx'
	artifacts_dir: 'D:\JB Pharma internal\JB Pharma\Cardiolek\Output\logger'
	```
2) Run:

```powershell
\.\.venv\Scripts\python main.py
```
