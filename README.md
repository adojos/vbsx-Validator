[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) 
![GitHub release (latest by date)](https://img.shields.io/github/v/release/testoxide/vbsx-Validator)
![GitHub repo size](https://img.shields.io/github/repo-size/testoxide/vbsx-Validator) ![platform](https://img.shields.io/badge/platform-win--32%20%7C%20win--64-lightgrey)

# vbsx-Validator
Free desktop utility for simple XML/XSD validation (supports multiple XSD). Fastest XML validation of 'large sized' files (in-memory) without rendering overhead.
DOM Parser based XML / XSD validation built on MSXML6. Supports full (multiple) error parsing of a given XML.

Also supports Batch (Multiple XML Files) as a single operation. Validate hundreds of XML against a single or multple XSD as one batch operation. Generates verbose log file for all operations and output.

No-frills, lightweight yet powerful! Built on windows native technologies!
This utility has no dependency on third party compiler / interpreters / engines (e.g. java, nodejs. .NET or other such runtimes).

### Download

Please do not download from source code section as it may not be stable.

Instead use any latest stable release version, available for download from _[Releases](https://github.com/testoxide/vbsx-Validator/releases)_ section

### How To run

Simply double-click the main script file named 'VBSX_Main.vbs' to launch the utility. This will launch the command line interface.
Please note that you might get UAC prompt if UAC is enabled on your windows.

When running under 'Bulk File Mode', the application would auto-create a main output folder and two sub-folders (valid file folder and Invalid file folder) on your filesystem. The app would auto save the XMLs being validated appropriately into these folders based on validation result.

**Note :** _Drag and Drop may work with UAC disabled._

_Refer [Wiki](https://github.com/testoxide/vbsx-Validator/wiki) for usage tips, screenshots and [video demo](https://github.com/testoxide/vbsx-Validator/wiki/Video-Demo-&-Overview)_


### Prerequisites

Win 7 / 8 / Server.

MSXML6 Official Download [here](https://www.microsoft.com/en-us/download/details.aspx?id=3988)

Require admin privileges / script execution privileges (elevated UAC prompt) on your windows system.


### Technical Notes (Design)

* The Validation Parser, accepts _multiple XSDs against a given XML_ (XMLs referencing multiple XSDs) for validation since version v2.0.1 (Multi XSD support). However the parser does not auto-import any schemas from network or filesystem, but rather considers only user supplied XSDs, hence make sure you provide all required XSDs.

* The Parser has been designed _'not to resolve externals'_. It does not evaluate or resolve the 'schemaLocation' or other attributes specified in DocumentRoot for locating schemas. The reason is that most of the time schemaLocation is not always valid or resolvable as XML travels system to system. _Hence this design avoids non-schema related errors_.

* The parser _validates strictly against the supplied XSD_ (schema definition) only without auto-resolving schemaLocation or other nameSpace attributes from the XML document. This provides robust validation against supplied XSD.

* The validation parser inherently validates all XML for well-formedness / structural.

* The validation _parser needs Namespace (targetNamespace)_ which is currently _extracted from the supplied XSD_. Please make sure that 'targetNamespace' declaration if any, in your XSD is correct. The _targetNamespace decalaration is not mandatory_ and hence XSD without targetNamespace are also validated properly.

* The current version does not support XML files with inline schemas

* Please refer ['Further Reading'](https://github.com/testoxide/vbsx-Validator/wiki/Additional-Notes) section of Wiki for more information if required.


### Built With

* [VBScript](https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/)
* [WScript](https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2003/cc738350(v=ws.10)) 
* [MSXML6](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms763742(v%3dvs.85))


### Known Issues / Bugs

Please refer the [Issue list](https://github.com/testoxide/vbsx-Validator/issues).

Feel free to contribte by logging any new defects, issues or enhancement requests

### Authors

* **Tushar Sharma**


### License

This is licensed under the MIT License - see the [LICENSE.md](https://github.com/testoxide/vbsx-Validator/blob/master/LICENSE) file for details

