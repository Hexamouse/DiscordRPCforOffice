# DiscordRPCforOffice

<!--
<p align="center">
  <img src="images/preview.png" alt="DiscordRPCforOffice Preview" width="300">
</p>
-->

**DiscordRPCforOffice** is an experimental C++ utility that synchronizes your Microsoft Office activity with Discord Rich Presence in real-time. It supports Word, Excel, PowerPoint, and Outlook across a wide Office compatibility range from Office 2007 up to the latest 2025-era Microsoft 365 releases.

> **EXPERIMENTAL**  
> Use at your own risk — this application interacts with Windows COM and active Microsoft Office processes.

## Key Features

- **Multi-App Office Detection**  
  Tracks active activity from Microsoft Word, Excel, PowerPoint, and Outlook automatically.

- **Broad Version Support**  
  Designed for Microsoft Office versions from 2007 through modern Microsoft 365 builds released up to 2025.

- **English Status Display**  
  Shows clean and professional activity information such as:

  ```text
  Editing Thesis.docx
  Page 3 of 15
  ```

- **Late Binding COM**  
  Uses `IDispatch` for version-independent Office object model access.

- **Performance Optimized**  
  Built with efficient C++17 implementation for low CPU and memory usage.

- **x64 Target**  
  Intended for modern Windows systems and current development toolchains.

***

## Supported Applications

| Application | Activity Example |
|---|---|
| Word | Editing Thesis.docx |
| Excel | Working on Budget.xlsx |
| PowerPoint | Presenting Seminar.pptx |
| Outlook | Reading Inbox or composing email |

***

## Technical Specs

| Aspect | Details |
|---|---|
| Language | C++17 |
| Core Tech | Discord RPC SDK, Windows COM |
| IDE | Visual Studio 2022 |
| Platform | x64 |
| Runtime Library | Multi-threaded DLL (/MD) |
| Office Scope | Word, Excel, PowerPoint, Outlook |
| Version Coverage | Office 2007 to Microsoft 365 (2025) |

***

## Dependencies

- **Discord RPC SDK**
  - Place headers in `deps/include`
  - Place `.lib` files in `deps/lib`

- **Windows SDK**
  - Required for COM integration and UTF-8 support
  - Included with Visual Studio 2022

- **Microsoft Office**
  - Compatible target scope includes Office 2007, 2010, 2013, 2016, 2019, 2021, 2024, and Microsoft 365

***

## Build Instructions

1. Clone the repository:

```bash
git clone https://github.com/Hexamouse/DiscordRPCforOffice.git
```

2. Open:

```text
DiscordRPCforOffice.sln
```

in Visual Studio 2022.

3. Select configuration:

```text
Release | x64
```

4. Navigate to:

```text
Project Properties
→ C/C++
→ Code Generation
→ Runtime Library
```

Set:

```text
Multi-threaded DLL (/MD)
```

5. Build the solution:

```text
Ctrl + Shift + B
```

***

## Usage

1. Launch one of the supported Microsoft Office applications.
2. Open or interact with a supported Office file or Outlook window.
3. Run:

```text
DiscordRPCforOffice.exe
```

4. Your Discord Rich Presence will automatically update based on the active Office application, such as:
   - Application name
   - File or window context
   - Current page, sheet, slide, or mail activity

### Example

```text
Editing Thesis.docx
Page 3 of 15
```

***

<!--
## Preview Image Setup

Preview image code is intentionally disabled for now. When needed later, place the image here:

```text
images/preview.png
```
-->

## Project Structure

```text
DiscordRPCforOffice/
├── include
│   ├── Core/
│   ├── Detectors/
│   ├── Helpers/
├── src/
│   ├── Core/
│   ├── Detectors/
│   ├── Helpers/
├── deps/
│   ├── include/
│   └── lib/
│
├── README.md
└── DiscordRPCforOffice.sln
```

***

## Compatibility Notes

- Office 2007 support may require more careful testing because older releases use legacy environments.
- Modern Microsoft 365 builds may change internal behavior more frequently than perpetual Office releases.
- Late Binding helps reduce dependency on a single Office version implementation.

***
