/*!-----------------------------------------------------------------------
	ids.h

	Contains all the progids/guids/clsids/strings used across the addin. If an
	id is needed to be updated then it should only need to be changed
	here and nowhere else.

	Remember these ids need to be unique on each system so they should be
	changed if this sample addin is used as a base for a another addin.

-----------------------------------------------------------------------!*/
#pragma once

// Strings
#define ADDIN_PROGID            OLESTR("KeepOutlookRunningCOMAddin.Connect")

// GUIDs
#define TYPELIB_GUID            870E8189-BD89-46B7-B3E3-25A79EB78459
#define ADDIN_CLSID             AA9CDF1C-A41E-43C4-81C8-7330D6A5322E

// Remember to keep these string version in sync with the ones above
#define TYPELIB_GUID_STR        OLESTR("870E8189-BD89-46B7-B3E3-25A79EB78459")
#define ADDIN_CLSID_STR         OLESTR("AA9CDF1C-A41E-43C4-81C8-7330D6A5322E")

