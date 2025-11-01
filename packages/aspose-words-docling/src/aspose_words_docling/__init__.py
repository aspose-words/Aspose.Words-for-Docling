# SPDX-FileCopyrightText: 2025-Aspose Pty Ltd
# SPDX-License-Identifier: MIT

from ._plugin import (
    __plugin_interface_version__, 
    register_converters, 
    AsposeDocumentConverter,
    LicenseManager
)

from .aspose_words_converter import (
    AsposeWordsConverter, 
    ResourceResolver
)


from .__about__ import __version__

__all__ = [
    "__version__",
    "__plugin_interface_version__",
    "register_converters",
    "AsposeDocumentConverter",
    "LicenseManager",
    "AsposeWordsConverter",
    "ResourceResolver"
]
