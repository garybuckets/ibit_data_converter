# iShares Bitcoin Trust ETF Data Downloader

This PowerShell script downloads the iShares Bitcoin Trust ETF fund data from the official iShares website and converts it from XLS format to XLSX format.

## Features

- Downloads fund data directly from iShares.com
- Converts XLS format to modern XLSX format
- Colored console output for better user experience
- Error handling and cleanup
- Optional parameters for customization

## Prerequisites

- **Windows PowerShell 5.1 or PowerShell Core 6.0+**
- **Microsoft Excel** (for XLS to XLSX conversion)
- Internet connection

## Usage

### Basic Usage
```powershell
.\download_ishares_bitcoin_etf.ps1
```
This will download the file and save it as `ishares_bitcoin_etf.xlsx` in the current directory.

### Advanced Usage

#### Specify custom output path:
```powershell
.\download_ishares_bitcoin_etf.ps1 -OutputPath "C:\Data\bitcoin_etf_data.xlsx"
```

#### Force overwrite existing file:
```powershell
.\download_ishares_bitcoin_etf.ps1 -Force
```

#### Combine parameters:
```powershell
.\download_ishares_bitcoin_etf.ps1 -OutputPath "C:\Data\bitcoin_etf_data.xlsx" -Force
```

## Parameters

- `-OutputPath`: Specifies the output file path (default: `.\ishares_bitcoin_etf.xlsx`)
- `-Force`: Forces overwrite of existing files without prompting

## What the Script Does

1. **Downloads** the iShares Bitcoin Trust ETF fund data from:
   ```
   https://www.ishares.com/us/products/333011/fund/1521942788811.ajax?fileType=xls&fileName=iShares-Bitcoin-Trust-ETF_fund&dataType=fund
   ```

2. **Converts** the downloaded XLS file to XLSX format using Excel COM automation

3. **Cleans up** temporary files automatically

4. **Provides** detailed progress information with colored output

## Error Handling

The script includes comprehensive error handling for:
- Network connectivity issues
- File download failures
- Excel availability checks
- File conversion errors
- Permission issues

## Output

- **Success**: Creates an XLSX file with the fund data
- **Failure**: Provides detailed error messages and cleanup

## Troubleshooting

### Excel Not Available
If Excel is not installed or accessible, the script will:
- Download the XLS file to a temporary location
- Inform you where the file is located
- Suggest manual conversion

### Permission Issues
Run PowerShell as Administrator if you encounter permission errors.

### Execution Policy
If you get execution policy errors, run:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## File Information

The downloaded file contains iShares Bitcoin Trust ETF fund data including:
- Fund performance metrics
- Holdings information
- Risk statistics
- Other fund-related data

## Security Notes

- The script uses a standard User-Agent header for web requests
- Temporary files are automatically cleaned up
- No sensitive data is logged or stored

## Version

This script was designed with help from Cursor/Claude Sonnet and is designed for Windows PowerShell environments. 
