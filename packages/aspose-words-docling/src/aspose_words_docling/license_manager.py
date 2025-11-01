import os, logging

import aspose.words as aw


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
