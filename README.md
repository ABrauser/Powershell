# 🛠️ PowerShell Tools

A collection of useful PowerShell scripts and functions by [Benjamin Rauser](mailto:Benjamin.Rauser@outlook.com).

## 📦 Tools

| Tool | Description |
|------|-------------|
| [DocPipeline](./DocPipeline/) | Document processing pipeline: Scan → Copy → Convert (Docling) → Upload (AnythingLLM) with interactive HTML dashboard |

## 🚀 Usage

Each tool lives in its own subfolder with a dedicated `README.md`. Click the links above for detailed documentation.

### Quick Example

```powershell
# Load and run the folder scanner
. .\DocPipeline\Invoke-FolderScan.ps1
Invoke-FolderScan -Path "C:\MyData" -Recurse

# Open the AnythingLLM upload GUI
. .\DocPipeline\Invoke-AnythingLLMUpload.ps1
Invoke-AnythingLLMUpload -Gui
```

## 📄 License

[MIT](LICENSE)
