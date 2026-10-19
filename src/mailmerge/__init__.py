"""
Performs a Mail Merge on Office Open XML (docx) files. Can be used on any
system without having to install Microsoft Office Word. Supports Python 3.7 and up.
"""

__version__ = "1.0.3"
__author__ = "Iulian Ciorăscu"
__credits__ = ["Iulian Ciorăscu", "Bouke Haarsma"]
__license__ = "MIT"
__maintainer__ = "Iulian Ciorăscu"
__email__ = "iulian.ciorascu@gmail.com"
__status__ = "Production"

from .constants import NAMESPACES as NAMESPACES
from .mailmerge import MailMerge as MailMerge
from .mergeoptions import MailMergeOptions as MailMergeOptions
from .mergeoptions import OptionAutoUpdateFields as OptionAutoUpdateFields
from .mergeoptions import OptionKeepFields as OptionKeepFields

__all__ = ["NAMESPACES", "MailMerge", "MailMergeOptions", "OptionAutoUpdateFields", "OptionKeepFields", "__version__"]
