Attribute VB_Name = "TAPIDecs"
'******************************************************************
'* Module:  TAPIDecs.bas
'* Purpose: Necessary TAPI APIs and UDTs
'*
'* Author:  Sumanta Ray
'* Email:   chandsumant@hotmail.com
'* Date:    14/07/2005
'* Copyright (c) 2003-2005 Sumanta Ray. All rights reserved
'******************************************************************
'* Note: Probably some portion was taken from Ray Mercer's code.
'*       Don't remeber. Anyway credit goes to them
'******************************************************************

Option Explicit

Type LINEINITIALIZEEXPARAMS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    dwOptions As Long
    hEvent As Long
    dwCompletionKey As Long
End Type

Type LINEDIALPARAMS
    dwDialPause As Long
    dwDialSpeed As Long
    dwDigitDuration As Long
    dwWaitForDialtone As Long
End Type

Type LINEEXTENSIONID
    dwExtensionID0 As Long
    dwExtensionID1 As Long
    dwExtensionID2 As Long
    dwExtensionID3 As Long
End Type

Type LINEDEVCAPS
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    dwProviderInfoSize As Long
    dwProviderInfoOffset As Long
    dwSwitchInfoSize As Long
    dwSwitchInfoOffset As Long
    dwPermanentLineID As Long
    dwLineNameSize As Long
    dwLineNameOffset As Long
    dwStringFormat As Long
    dwAddressModes As Long
    dwNumAddresses As Long
    dwBearerModes As Long
    dwMaxRate As Long
    dwMediaModes As Long
    dwGenerateToneModes As Long
    dwGenerateToneMaxNumFreq As Long
    dwGenerateDigitModes As Long
    dwMonitorToneMaxNumFreq As Long
    dwMonitorToneMaxNumEntries As Long
    dwMonitorDigitModes As Long
    dwGatherDigitsMinTimeout As Long
    dwGatherDigitsMaxTimeout As Long
    dwMedCtlDigitMaxListSize As Long
    dwMedCtlMediaMaxListSize As Long
    dwMedCtlToneMaxListSize As Long
    dwMedCtlCallStateMaxListSize As Long
    dwDevCapFlags As Long
    dwMaxNumActiveCalls As Long
    dwAnswerMode As Long
    dwRingModes As Long
    'dwLineStates As Long
    dwLineStates As Long
    dwUUIAcceptSize As Long
    dwUUIAnswerSize As Long
    dwUUIMakeCallSize As Long
    dwUUIDropSize As Long
    dwUUISendUserUserInfoSize As Long
    dwUUICallInfoSize As Long
    MinDialParams As LINEDIALPARAMS
    MaxDialParams As LINEDIALPARAMS
    DefaultDialParams As LINEDIALPARAMS
    dwNumTerminals As Long
    dwTerminalCapsSize As Long
    dwTerminalCapsOffset As Long
    dwTerminalTextEntrySize As Long
    dwTerminalTextSize As Long
    dwTerminalTextOffset As Long
    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    dwLineFeatures As Long                                 '// TAPI v1.4
'#if (TAPI_CURRENT_VERSION >= 0x00020000)
    dwSettableDevStatus As Long                            '// TAPI v2.0
    dwDeviceClassesSize As Long                            ' // TAPI v2.0
    dwDeviceClassesOffset As Long                          ' // TAPI v2.0
'#End If
    vbByteBuffer(0 To 2048) As Byte
    'note*  if you get LINEERR_STRUCTURETOOSMALL and you know that you are
    'doing everything else right (Like initializing the dwActualSize parameter
    'of structs you are passing) then you *might* need to increase this buffer
    'size and recompile.  However, I have not had any problems with this size yet.
End Type

Public Type LINECALLPARAMS                 '// DEFAULTS
    dwTotalSize As Long                    '// ---------
    dwBearerMode As Long                   '// voice
    dwMinRate As Long                      '// (3.1kHz)
    dwMaxRate As Long                      '// (3.1kHz)
    dwMediaMode As Long                    '// interactiveVoice
    dwCallParamFlags As Long               '// 0
    dwAddressMode As Long                  '// addressID
    dwAddressID As Long                    '// (any available)
    DialParams As LINEDIALPARAMS           '// (0, 0, 0, 0)
    dwOrigAddressSize As Long              '// 0
    dwOrigAddressOffset As Long
    dwDisplayableAddressSize As Long
    dwDisplayableAddressOffset As Long
    dwCalledPartySize As Long              '// 0
    dwCalledPartyOffset As Long
    dwCommentSize As Long                  '// 0
    dwCommentOffset As Long
    dwUserUserInfoSize As Long             '// 0
    dwUserUserInfoOffset As Long
    dwHighLevelCompSize As Long            '// 0
    dwHighLevelCompOffset As Long
    dwLowLevelCompSize As Long             '// 0
    dwLowLevelCompOffset As Long
    dwDevSpecificSize As Long              '// 0
    dwDevSpecificOffset As Long
'#if (TAPI_CURRENT_VERSION >= 0x00020000)
    dwPredictiveAutoTransferStates As Long                 '// TAPI v2.0
    dwTargetAddressSize As Long                            '// TAPI v2.0
    dwTargetAddressOffset As Long                          '// TAPI v2.0
    dwSendingFlowspecSize As Long                          '// TAPI v2.0
    dwSendingFlowspecOffset As Long                        '// TAPI v2.0
    dwReceivingFlowspecSize As Long                        '// TAPI v2.0
    dwReceivingFlowspecOffset As Long                      '// TAPI v2.0
    dwDeviceClassSize As Long                              '// TAPI v2.0
    dwDeviceClassOffset As Long                            '// TAPI v2.0
    dwDeviceConfigSize As Long                             '// TAPI v2.0
    dwDeviceConfigOffset As Long                           '// TAPI v2.0
    dwCallDataSize As Long                                 '// TAPI v2.0
    dwCallDataOffset As Long                               '// TAPI v2.0
    dwNoAnswerTimeout As Long                              '// TAPI v2.0
    dwCallingPartyIDSize As Long                           '// TAPI v2.0
    dwCallingPartyIDOffset As Long                         '// TAPI v2.0
'#End If
End Type


Type LINECALLINFO
    
    dwTotalSize As Long
    dwNeededSize As Long
    dwUsedSize As Long
    hLine As Long
    dwLineDeviceID As Long
    dwAddressID As Long
    dwBearerMode As Long
    dwRate As Long
    dwMediaMode As Long
    dwAppSpecific As Long
    dwCallID As Long
    dwRelatedCallID As Long
    dwCallParamFlags As Long
    dwCallStates As Long
    dwMonitorDigitModes As Long
    dwMonitorMediaModes As Long
    DialParams As LINEDIALPARAMS
    dwOrigin As Long
    dwReason As Long
    dwCompletionID As Long
    dwNumOwners As Long
    dwNumMonitors As Long
    dwCountryCode As Long
    dwTrunk As Long
    dwCallerIDFlags As Long
    dwCallerIDSize As Long
    dwCallerIDOffset As Long
    dwCallerIDNameSize As Long
    dwCallerIDNameOffset As Long
    dwCalledIDFlags As Long
    dwCalledIDSize As Long
    dwCalledIDOffset As Long
    dwCalledIDNameSize As Long
    dwCalledIDNameOffset As Long
    dwConnectedIDFlags As Long
    dwConnectedIDSize As Long
    dwConnectedIDOffset As Long
    dwConnectedIDNameSize As Long
    dwConnectedIDNameOffset As Long
    dwRedirectionIDFlags As Long
    dwRedirectionIDSize As Long
    dwRedirectionIDOffset As Long
    dwRedirectionIDNameSize As Long
    dwRedirectionIDNameOffset As Long
    dwRedirectingIDFlags As Long
    dwRedirectingIDSize As Long
    dwRedirectingIDOffset As Long
    dwRedirectingIDNameSize As Long
    dwRedirectingIDNameOffset As Long
    dwAppNameSize As Long
    dwAppNameOffset As Long
    dwDisplayableAddressSize As Long
    dwDisplayableAddressOffset As Long
    dwCalledPartySize As Long
    dwCalledPartyOffset As Long
    dwCommentSize As Long
    dwCommentOffset As Long
    dwDisplaySize As Long
    dwDisplayOffset As Long
    dwUserUserInfoSize As Long
    dwUserUserInfoOffset As Long
    dwHighLevelCompSize As Long
    dwHighLevelCompOffset As Long
    dwLowLevelCompSize As Long
    dwLowLevelCompOffset As Long
    dwChargingInfoSize As Long
    dwChargingInfoOffset As Long
    dwTerminalModesSize As Long
    dwTerminalModesOffset As Long
    dwDevSpecificSize As Long
    dwDevSpecificOffset As Long
    bBytes(2000) As Byte 'HACK Added to TAPI structure for callinfo data.

End Type

'******************************************************************************************************

Declare Function lineInitialize Lib "TAPI32.DLL" _
    (ByRef lphLineApp As Long, _
    ByVal hInstance As Long, _
    ByVal lpfnCallback As Long, _
    ByVal lpszAppName As String, _
    ByRef lpdwNumDevs As Long) As Long

Declare Function lineInitializeEx Lib "TAPI32.DLL" Alias "lineInitializeExA" _
    (ByRef lphLineApp As Long, _
    ByVal hInstance As Long, _
    ByVal lpfnCallback As Long, _
    ByVal lpszFriendlyAppName As String, _
    ByRef lpdwNumDevs As Long, _
    ByRef lpdwAPIVersion As Long, _
    ByRef lpLineInitializeExParams As LINEINITIALIZEEXPARAMS) As Long

Declare Function lineGetDevCaps Lib "TAPI32.DLL" Alias "lineGetDevCapsA" _
    (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByVal dwAPIVersion As Long, _
    ByVal dwExtVersion As Long, _
    ByRef lpLineDevCaps As LINEDEVCAPS) As Long
 
Declare Function lineGetCallInfo Lib "Tapi32" (ByVal hCall As Long, _
    ByRef lpCallInf As LINECALLINFO) As Long

Declare Function lineDeallocateCall Lib "TAPI32.DLL" _
    (ByVal hCall As Long) As Long

Declare Function lineNegotiateExtVersion Lib "TAPI32.DLL" (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByVal dwAPIVersion As Long, _
    ByVal dwExtLowVersion As Long, _
    ByVal dwExtHighVersion As Long, _
    ByVal lpdwExtVersion As Long)
    
Declare Function lineNegotiateAPIVersion Lib "TAPI32.DLL" _
    (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByVal dwAPILowVersion As Long, _
    ByVal dwAPIHighVersion As Long, _
    ByRef lpdwAPIVersion As Long, _
    ByRef lpExtensionID As LINEEXTENSIONID) As Long

Declare Function lineClose Lib "TAPI32.DLL" _
    (ByVal hLine As Long) As Long

Declare Function lineSetStatusMessages Lib "TAPI32.DLL" (ByVal hLine As Long, ByVal dwLineStates As Long, ByVal dwAddressStates As Long) As Long

Declare Function lineShutdown Lib "TAPI32.DLL" _
    (ByVal hLineApp As Long) As Long

Declare Function lineOpen Lib "TAPI32.DLL" _
    (ByVal hLineApp As Long, _
    ByVal dwDeviceID As Long, _
    ByRef lphLine As Long, _
    ByVal dwAPIVersion As Long, _
    ByVal dwExtVersion As Long, _
    ByVal dwCallbackInstance As Long, _
    ByVal dwPrivileges As Long, _
    ByVal dwMediaModes As Long, _
    ByRef lpCallParams As Any) As Long






