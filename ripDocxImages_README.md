# ripDocxImages.ps1

A simple Windows PowerShell script to extract all images from a Microsoft Word `.docx` file.

---

## Features

- Extracts all embedded images from a `.docx` file.
- Creates a dedicated output folder named after the source file.
- Handles missing or empty documents gracefully.
- Cleans up temporary extraction files automatically.

---

## Requirements

- Windows PowerShell 5.0 or higher.
- A valid `.docx` file.

---

## Usage

```powershell
.\ripDocxImages.ps1 -Path "C:\path\to\your\document.docx"
```

### Example

```powershell
.\ripDocxImages.ps1 -Path "C:\docs\MyDocument.docx"
```

This will:

1. Validate the input file.
2. Create a folder `C:\docs\MyDocument_images`.
3. Extract all images from `MyDocument.docx` into that folder.
4. Clean up any temporary files used during extraction.

---

## Parameters

| Parameter | Description | Required |
|-----------|-------------|----------|
| `-Path`   | Full path to the `.docx` file. | Yes |

---

## Notes

- If no images are found, the output folder will be removed.
- Only `.docx` files are supported. `.doc` or other formats will trigger an error.
- The script uses a temporary directory for extraction and automatically cleans it up.
- **Supported image types:** .jpg, .jpeg, .png, .gif, .bmp, .tiff (depends on document).

---

## License

This script is released under the MIT License. Feel free to use, modify, and share.
