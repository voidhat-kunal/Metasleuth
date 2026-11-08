# MetaSleuth
### Forensic Metadata Analysis Tool

**Inspect Digital Evidence. Examine Metadata. Preserve Investigative Context.**

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white" alt="Platform: Windows">
  <img src="https://img.shields.io/badge/Language-Python-3776AB?logo=python&logoColor=white" alt="Language: Python">
  <img src="https://img.shields.io/badge/Version-1.0-2563EB" alt="Version 1.0">
  <img src="https://img.shields.io/badge/Status-Development-orange" alt="Development Status">
</p>

---

## Overview

**MetaSleuth** is a desktop-oriented digital forensics utility designed to examine digital evidence and analyze file-related forensic information through a unified interface.

It brings together metadata inspection, cryptographic hash analysis, hexadecimal viewing, strings and indicator extraction, image forensics, timeline analysis, file comparison, and audit logging in one application.

Built with Python, MetaSleuth aims to simplify evidence examination and provide investigators with a structured environment for reviewing digital artifacts.

> **Important:** Application-generated labels such as `CLEAN` or `LOW` are analytical results, not guarantees that a file is safe or free from malicious content. Findings should always be validated using appropriate forensic procedures and independent tools.

---

## Core Features

### 1. Evidence Management
- Import individual files and folders.
- Review loaded evidence in a centralized list.
- Remove selected evidence or clear the current list.

### 2. Metadata Analysis
- Examine file-system metadata.
- Inspect file signatures and identifying information.
- Review available hash and media-related metadata.

### 3. Cryptographic Hash Analysis
- Generate MD5, SHA-1, SHA-256, and SHA-512 hashes.
- Compare calculated hashes with known values.
- Support file integrity verification workflows.

### 4. Hexadecimal Viewer
- Inspect raw file data in hexadecimal format.
- View corresponding ASCII representations.
- Navigate and search through file contents.

### 5. Strings & IOC Extraction
- Extract printable strings from digital files.
- Identify URLs, email addresses, and IP addresses.
- Discover file paths and other potentially relevant indicators.

### 6. Image Forensics
- Preview image evidence.
- Perform Error Level Analysis (ELA).
- Examine image-related forensic information.

### 7. Timeline Analysis
- Review available file and metadata events.
- Organize relevant timestamps for investigation.
- Export timeline information in CSV format.

### 8. File Comparison
- Compare properties of two evidence items.
- Identify differences between selected files.

### 9. Audit Logging
- Review recorded application activities.
- Maintain an investigative activity record.
- Export audit log information.

### 10. Forensic Reporting
- Generate forensic reports from the application.
- Organize relevant examination findings for documentation.

---

## Technology

| Component | Technology |
|---|---|
| Programming Language | Python |
| Application Type | Desktop Application |
| Operating System Target | Windows |
| Analysis Focus | Digital Forensics & Metadata Analysis |

---

## Getting Started

### Prerequisites

- Windows 10 or Windows 11
- Python 3.10 or newer (subject to actual source-code compatibility)
- Git

### Installation

**1. Clone the repository**

```bash
git clone https://github.com/voidhat-kunal/MetaSleuth.git
```

**2. Navigate to the project directory**

```bash
cd MetaSleuth
```

**3. Create a virtual environment**

```bash
python -m venv .venv
```

**4. Activate the environment**

```bash
.venv\Scripts\activate
```

**5. Install dependencies**

After the project's verified dependencies have been added to `requirements.txt`:

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

**6. Launch MetaSleuth**

```bash
python Metasleuth.py
```

---

## Digital Forensics & Responsible Use

MetaSleuth is intended to support authorized digital evidence examination and forensic analysis.

Users should follow responsible forensic practices:

- Examine only files they own or are authorized to investigate.
- Preserve original evidence and work with verified copies whenever possible.
- Maintain appropriate evidence documentation and chain-of-custody records.
- Record hash values when integrity verification is required.
- Validate metadata, timestamps, extracted indicators, and analytical findings.
- Avoid treating automated analysis results as conclusive proof.

---

## Project Roadmap

- [ ] Expand metadata analysis capabilities.
- [ ] Improve supported file-format coverage.
- [ ] Enhance IOC extraction and analysis.
- [ ] Improve forensic report generation.
- [ ] Strengthen evidence integrity and audit workflows.
- [ ] Perform additional testing and validation.
- [ ] Prepare stable versioned releases.

---

## Contributing

Contributions, suggestions, and bug reports are welcome.

Please review the project's contribution guidelines before submitting changes.

## Security

If you discover a security vulnerability, please follow the responsible disclosure instructions in `SECURITY.md`.

## License

**All Rights Reserved.**

No permission to use, copy, modify, or distribute this software is granted without prior written authorization from the copyright holder. Please refer to the `LICENSE` file for details.

---

## Author

**Voidhat Kunal**

Digital Forensics | Cybersecurity | Ethical Hacking | Cyber Crime Investigation

- GitHub: [@voidhat-kunal](https://github.com/voidhat-kunal)

---

<p align="center">

**MetaSleuth — Investigate. Analyze. Discover.**

*Bringing structure and clarity to digital evidence examination.*

</p>
