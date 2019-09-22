[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) 
![GitHub release (latest by date)](https://img.shields.io/github/v/release/testoxide/vbsx-Validator)
![GitHub repo size](https://img.shields.io/github/repo-size/testoxide/vbsx-Validator)

# vbsx-Validator
Free script utility for XML/XSD validation. Fastest XML validation of 'large sized' files (in-memory) without rendering overhead.
DOM Parser based XML / XSD validation built on MSXML6. Supports full (Multiple) error parsing of a given XML.

Also supports Batch (Multiple XML Files) as a single operation. Validate hundreds of XML against a given XSD as one batch operation. Generates verbose log file for all operations and output.

No-frills, lightweight yet powerful! Built on windows native technologies!
This utility has no dependency on third party compiler / interpreters / engines (e.g. java, nodejs. .NET or other such runtimes).

### Download

Please do not download from source code section as it may not be stable.

Instead use any latest stable release version, available for download from _[Releases](https://github.com/testoxide/vbsx-Validator/releases)_ section

### How To run

Simply double-click the main script file named 'VBSX_Main.vbs' to launch the utility. This will launch the command line interface.
Please note that you might get UAC prompt if UAC is enabled on your windows.

**Note :** _Drag and Drop may work with UAC disabled._

_Refer [Wiki](https://github.com/testoxide/vbsx-Validator/wiki) for usage tips, screenshots and [video demo](https://www.youtube.com/watch?v=bjuY4CBv5iM)_


### Prerequisites

Win 7 / 8 / Server.

Require admin privileges / script execution privileges (elevated UAC prompt) on your windows system.


### Technical Notes (Design)

* The Parser has been designed _not to resolve externals_. It does not evaluate or resolve the schemaLocation or attributes specified in DocumentRoot. The reason is that most of the time schemaLocation is not always valid or resolvable. Hence this design avoids non-schema related errors.

* The parser validates _strictly against the supplied XSD (schema definition) only without auto-resolving schemaLocation_ or other nameSpace attributes.

* The parser needs Namespace (targetNamespace) which is currently extracted from the supplied XSD.

* The current version does not support XML files with inline schemas


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

