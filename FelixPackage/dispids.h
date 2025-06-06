
#ifndef Z80_PACKAGE_DISPIDS_H_INCLDUDED
#define Z80_PACKAGE_DISPIDS_H_INCLDUDED

// Common to all XxxProperties interfaces.
// Visual Studio wants a property named "__name" to show it at the top of the Properties Window.
// Our XML saving code ignores properties with the dispid equal to DISPID_VALUE.
#define dispid__name DISPID_VALUE

#define dispidLoadAddress          2
#define dispidEntryPointAddress    3
#define dispidLaunchType           4
#define dispidConfigurations       5
#define dispidAssemblerProperties  6
#define dispidDebuggingProperties  7
#define dispidItems                8
#define dispidProjectGuid          9
#define dispidSaveListing          10
#define dispidListingFilename      11
#define dispidProjectDir           12
#define dispidOutputFileType       13
#define dispidPath                 14
#define dispidBuildToolKind        15

#endif