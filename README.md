# Aspose.Words for Docling

![Python Version](https://img.shields.io/badge/python-3.12+-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## Overview
**Aspose.Words for Docling** is a free plugin for [Docling](https://github.com/docling-project/docling) based on [Aspose.Words for Python via .Net](https://products.aspose.com/words/python-net/) commercial library.  
The plugin is designed for parsing multiple document formats and converting them into Docling Documents suitable for AI processing.  
**Aspose.Words** plugin allows you to convert formats as `.docx`, `.pdf`, `.doc`, `.rtf`, `.html`, `.mhtml`, `.mobi`, `.azw3`, `.epub`, `.odt`, `.txt`, `.md` and `.xml`.


## Features

- Convert `.docx`, `.pdf`, `.doc`, `.rtf`, `.html`, `.mhtml`, `.mobi`, `.azw3`, `.epub`, `.odt`, `.txt`, `.md` and `.xml` to [DoclingDocument](https://docling-project.github.io/docling/concepts/docling_document/).
- Support all document components, including paragraphs, tables, images, headers, and footers.

## Requirements

- [Docling](https://github.com/docling-project/docling) version 2.60.0 or higher.
- [Aspose.Words for Python via .Net](https://products.aspose.com/words/python-net/). This library is a [commercial product](https://purchase.aspose.com/buy/words/python).  
You'll need to obtain a valid license for Aspose.Words. The package will install this dependency, but you're responsible for complying with Aspose's licensing terms.

## Installation

```bash
pip install aspose-words-docling
```

### Python API

```python
from aspose_words_docling import AsposeWordsConverter

aspose_converter = AsposeWordsConverter()
doc = aspose_converter.convert("test.doc")

md = doc.export_to_markdown()
print(md)
```


## Set License

### Environment Variables
To activate your Aspose.Words for Python license, set the corresponding environment variable.  
Refer to the OS-specific instructions below:

**Unix-based (Linux/macOS):**
```
export ASPOSE_WORDS_LICENSE_PATH="/path/to/license/aspose.words.lic"
```
**Windows-based:**
```
set ASPOSE_WORDS_LICENSE_PATH=c:\path\to\license\aspose.words.lic
```
**Python API:**
```
from aspose_words_docling import LicenseManager

LicenseManager().apply_license("/path/to/license/aspose.words.lic")
```

## Running Tests

To run unit tests for **Aspose.Words for Docling**, follow these steps:

### 1. Navigate to the package directory

From the root of the repository, change into the package directory:

```bash
cd /packages/aspose-words-docling/tests
```

### 2. Install test dependencies

Make sure `pytest` is installed:

```bash
pip install pytest
```

### 3. Run tests using `pytest`

To run all tests:

```bash
pytest
```

## License

This package is licensed under the MIT License. However, it depends on Aspose.Words for Python via .Net library, which is proprietary, closed-source library.

⚠️ You must obtain valid license for Aspose.Words for Python via .Net library. This repository does not include or distribute any proprietary components.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of IBM
trademarks or logos is subject to and must follow
[IBM's Trademarks](https://www.ibm.com/legal/copyright-trademark).
[IBM Logo and Brand Guidelines](https://www.ibm.com/design/language/files/IBM_Logo_3rdParties_300822.pdf).
Use of IBM trademarks or logos in modified versions of this project must not cause confusion or imply IBM sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
