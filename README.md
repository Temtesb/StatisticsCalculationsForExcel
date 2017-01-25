# StatisticsCalculationsForExcel

These calculations, functions and subroutines are primarily focused on statistics have been built (or rebuilt from other open source/public domain libraries/projects) in Visual Basic for Applications (VBA) for use primarily in the Microsoft Office Suite. 

## Installation

Building from source code has some automation, but is still a couple of manual steps:

1. [Download](../../archive/master.zip) this repository from GitHub and unzip it
2. Create or move a excel or access file into either the root or a existing 'build' directory of the project and open that file
3. Manually import the build_ImportExport.bas module and run the 'ToolImportModules' subroutine 

  1. <kbd>alt</kbd> + <kbd>f11</kbd> (Launch IDE)
  2. <kbd>ctl</kbd> + <kbd>m</kbd> (Import File)
  3. <kbd>ctl</kbd> + <kbd>g</kbd> (Goto Imediate Window)
  4. Type and execute 'ToolImportModules' from the imediate window

4. Once the code is loaded you can save, close and move your file.

## Usage

TODO: Write usage instructions

## Contributing

1. Fork it!
2. Create your feature branch: `git checkout -b my-new-feature`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin my-new-feature`
5. Submit a pull request :D
6. Provide your comments to an existing [Issue](../../issues) or [submit a new issue](../../issues/new) if a simmilar one does not exist.

## History

This project began as a way to build vba functions to statisticaly analyze data using the software that was avialable to us in our work.

The initial functions where built during a Managing Small Business Data Analytics class at [Regent University](www.regent.edu/). Those functions that are mentioned as being derived from the course text of 'Essentials of Business Analytics ISBN 978-1-305-62773-4' are all found in the [/Excel/Src/basStatsistics.bas](Excel/src/bas_Statistics.bas) file or [General%20VB/src/VB_Statistics.bas](General VB/Src/VB_Statsistics.bas) file.

## Credits

This project has been created by (and is managed by) [William Young](mailto:wmyoung708@gmail.com) and [Jeremy D. Gerdes](mailto:jeremy.gerdes@navy.mil) 

Any additional specific attributions to individuals or other [copyleft](https://copyleft.org/) sources are listed in [NOTICE.md](NOTICE.md).

## License

When this project is distributed in it's entirety the license is GNU GPL 3.0.  Many functions and source files are from multiple projects and have multiple open source or have public source attributions (see [NOTICE.md](NOTICE.md)).

All contributions to this project will be published under GNU GPL 3.0. By submitting a pull request, you are certifying that all content added to this project is either in a (c) license that is [compatible](https://www.gnu.org/licenses/license-list.en.html#GPLCompatibleLicenses) with the GNU GPL 3.0 license or that you publish those contributions under the GNU GPL 3.0 license.

The full text of each licens is listed in [LICENSE.md](LICENSE.md)
