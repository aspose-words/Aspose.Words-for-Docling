# Copyright (c) 2001-2026 Aspose Pty Ltd.

from io import BytesIO
import json, os, logging
from docling_core.types.doc.document import DoclingDocument

import aspose.words as aw


class AsposeWordsConverter:
    def convert(self, file: str) -> DoclingDocument:
        license_manager = LicenseManager()
        license_manager.apply_license() 

        aspose_doc = aw.Document(file)
        aspose_doc.update_list_labels()
        buffer = BytesIO()
        save_options = aw.saving.DoclingSaveOptions()
        aspose_doc.save(buffer, save_options)
        doc_dict = json.loads(buffer.getvalue())
        origin = doc_dict.get("origin")
        if origin is not None:
            mimetype = origin.get("mimetype")
        # Docling doesn't support some mime types.
        if (mimetype == "application/rtf" or mimetype == "multipart/related"):
            origin["mimetype"] = "text/plain"

        docling_doc = DoclingDocument.model_validate(doc_dict)
        return docling_doc
    

""" To activate your Aspose License, set the corresponding environment variable.
    Refer to the OS-specific instructions below:
    Unix-based (Linux/macOS):
        export ASPOSE_WORDS_LICENSE_PATH="/path/to/license/aspose.words.lic"
    Windows-based:
        set ASPOSE_WORDS_LICENSE_PATH=c:\\path\\to\\license\\aspose.words.lic """
class LicenseManager:
    def apply_license(self, license_path: str = None):
        if license_path is None:
            license_path = os.getenv("ASPOSE_WORDS_LICENSE_PATH")
        if license_path is not None and os.path.exists(license_path):
            logging.info(f"Applying Aspose.Words license from: {license_path}")
            license = aw.License()
            license.set_license(license_path)
        else:
            logging.warning("No valid Aspose.Words license found. Running in Evaluation mode. Please set the ASPOSE_WORDS_LICENSE_PATH environment variable.")
