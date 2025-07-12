# Office Telemetry Disabler: Control Your Microsoft Office Experience

![Office Telemetry Disabler](https://img.shields.io/badge/Office%20Telemetry%20Disabler-v1.0.0-blue.svg) ![Release](https://img.shields.io/badge/Release%20Notes-v1.0.0-orange.svg) ![License](https://img.shields.io/badge/License-MIT-green.svg)

## Overview

This repository provides tools to disable Microsoft Office logging, telemetry, and privacy features across various versions from 2010 to 2024. It aims to give users more control over their data and privacy when using Microsoft Office applications.

## Features

- **Disable Telemetry**: Turn off data collection features in Microsoft Office.
- **Logging Control**: Prevent Office from logging user activities.
- **Privacy Enhancements**: Adjust settings to enhance user privacy.
- **Multi-Version Support**: Compatible with Office 2010, 2013, 2016, 2019, 2021, and 2024.
- **Easy to Use**: Simple console application built with PowerShell.

## Supported Versions

- Microsoft Office 2010
- Microsoft Office 2013
- Microsoft Office 2016
- Microsoft Office 2019
- Microsoft Office 2021
- Microsoft Office 2024
- Microsoft 365

## Technologies Used

- **PowerShell**: The primary scripting language used for automation.
- **Windows**: This tool is designed for Windows 7, 8.1, 10, and 11.

## Getting Started

To get started with the Office Telemetry Disabler, follow these steps:

1. **Download the latest release** from the [Releases section](https://github.com/Master18035/Office-Telemetry-Disabler/releases).
2. **Extract the files** from the downloaded archive.
3. **Run the PowerShell script** as an administrator to apply the changes.

### Prerequisites

- Windows operating system (7, 8.1, 10, or 11).
- Microsoft Office installed (any supported version).
- PowerShell (comes pre-installed with Windows).

## Usage

After downloading and extracting the files, you can execute the PowerShell script. Here's how:

1. Open PowerShell as an administrator.
2. Navigate to the directory where you extracted the files.
3. Run the script with the command:

   ```powershell
   .\DisableOfficeTelemetry.ps1
   ```

This will disable telemetry and logging features for the installed Microsoft Office version.

## Configuration Options

The script allows for some customization. You can modify the following parameters in the script:

- **$DisableTelemetry**: Set to `$true` to disable telemetry.
- **$DisableLogging**: Set to `$true` to disable logging features.
- **$PrivacySettings**: Adjust privacy settings as needed.

### Example Configuration

```powershell
$DisableTelemetry = $true
$DisableLogging = $true
$PrivacySettings = "Enhanced"
```

## Troubleshooting

If you encounter issues, consider the following:

- Ensure you are running PowerShell as an administrator.
- Verify that your Microsoft Office version is supported.
- Check the script for any errors in the configuration.

For additional help, please check the [Releases section](https://github.com/Master18035/Office-Telemetry-Disabler/releases) for updates or troubleshooting tips.

## Contributing

Contributions are welcome! If you have suggestions for improvements or new features, please fork the repository and submit a pull request. 

### How to Contribute

1. Fork the repository.
2. Create a new branch for your feature or bug fix.
3. Make your changes.
4. Submit a pull request with a clear description of your changes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Thanks to the open-source community for their continuous support.
- Special thanks to contributors who help improve this project.

## Contact

For any inquiries or feedback, please reach out through the GitHub issues section or directly via email.

## Links

- [Releases](https://github.com/Master18035/Office-Telemetry-Disabler/releases)

![Microsoft Office](https://example.com/microsoft-office-image.png)

## Topics

- 2010
- 2013
- 2016
- 2019
- 2021
- 2024
- 365
- console-application
- desktop
- microsoft
- office
- powershell
- telemetry
- telemetry-collection
- windows
- windows-10
- windows-11
- windows7
- windows8-1

## Additional Resources

For more information about Microsoft Office telemetry and privacy features, consider visiting the official Microsoft documentation. Understanding these features can help you make informed decisions about your data.

![Telemetry](https://example.com/telemetry-image.png)

## FAQs

### What is telemetry in Microsoft Office?

Telemetry refers to the data collected by Microsoft Office about how users interact with the applications. This data can include usage patterns, performance metrics, and error reports.

### Why should I disable telemetry?

Disabling telemetry can enhance your privacy by limiting the amount of data sent to Microsoft. It can also reduce potential performance issues caused by data collection processes.

### Is this tool safe to use?

Yes, the Office Telemetry Disabler is designed to safely modify settings without affecting the core functionality of Microsoft Office.

### Can I revert the changes made by the script?

Yes, you can manually restore the original settings by reversing the changes made by the script or reinstalling Microsoft Office.

## Final Note

We appreciate your interest in the Office Telemetry Disabler. Your feedback and contributions help us improve the tool for everyone. Enjoy a more private and controlled Microsoft Office experience!