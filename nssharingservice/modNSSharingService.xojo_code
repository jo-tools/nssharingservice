#tag Module
Protected Module modNSSharingService
	#tag CompatibilityFlags = (TargetDesktop and (Target32Bit or Target64Bit))
	#tag Method, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Private Function BuildNSSharingServiceDelegateClass() As Ptr
		  #If TargetMacOS And Target64Bit Then
		    Declare Function NSClassFromString Lib "Foundation" (aClassName As CFStringRef) As Ptr
		    Declare Function objc_allocateClassPair Lib "Foundation" (superclass As Ptr, name As CString, extraBytes As Integer) As Ptr
		    Declare Function objc_lookUpClass Lib "Foundation" (name As CString) As Ptr
		    Declare Sub objc_registerClassPair Lib "Foundation" (cls As Ptr)
		    Declare Function objc_getProtocol Lib "Foundation" (name As CString) As Ptr
		    Declare Function class_addProtocol Lib "Foundation" (cls As Ptr, proto As Ptr) As Boolean
		    Declare Function class_addMethod Lib "Foundation"  (cls As Ptr, name As Ptr, imp As Ptr, types As CString) As Boolean
		    Declare Function NSSelectorFromString Lib "Foundation" (aSelectorName As CFStringRef) As Ptr
		    Declare Function alloc Lib "Foundation" Selector "alloc" (classRef As Ptr) As Ptr
		    Declare Function init Lib "Foundation" Selector "init" (classRef As Ptr) As Ptr
		    
		    'create the NSSharingServiceDelegateHandler
		    Var delegateClass As Ptr = objc_allocateClassPair(NSClassFromString("NSObject"), "NSSharingServiceDelegateHandler", 0)
		    If (delegateClass = Nil) Then
		      'Maybe it already exists. Let's check:
		      delegateClass = objc_lookUpClass("NSSharingServiceDelegateHandler")
		    Else
		      'no, so let's register class pair
		      objc_registerClassPair(delegateClass)
		      
		      'and add the protocol: NSSharingServiceDelegate
		      Var delegateProtocol As Ptr = objc_getProtocol("NSSharingServiceDelegate")  //get protocol type
		      If (Not class_addProtocol(delegateClass, delegateProtocol)) Then Break //add protocol to class
		    End If
		    If (delegateClass = Nil) Then Break
		    
		    'add the Xojo-implementation of the class methods
		    If (Not class_addMethod(delegateClass, NSSelectorFromString("sharingService:didFailToShareItems:error:"), AddressOf Delegate_Implementation_DidFailToShare, "v@:@@@")) Then Break
		    If (Not class_addMethod(delegateClass, NSSelectorFromString("sharingService:didShareItems:"), AddressOf Delegate_Implementation_DidShareItems, "v@:@@")) Then Break
		    If (Not class_addMethod(delegateClass, NSSelectorFromString("sharingService:sourceWindowForShareItems:sharingContentScope:"), AddressOf Delegate_Implementation_SourceWindow, "@@:@@^l")) Then Break
		    
		    Return init(alloc(delegateClass))
		  #EndIf
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Function ComposeEmail(psTo As String, psSubject As String, psBody As String, poAttachments() As FolderItem, poShowWithin As DesktopWindow, poResultCallback As ResultCallbackDelegate) As Boolean
		  #If TargetMacOS And Target64Bit Then
		    'This will perform the requested Sharing Service:
		    '- items from the parameters
		    '- poShowWithin: attach to that window (modally); nil -> sharing window will be independent
		    '- poResultCallback: reports back once Sharing Service has finished (success | error)
		    'Returns true - if NSSharingService could be invoked (not necessarily: has successfully shared the items!)
		    '               then wait for the callback with success or error
		    'Return false - if NSSharingService can't be invoked
		    Return PerformWithItems(NSSharingServiceName.ComposeEmail, psTo, psSubject, psBody, poAttachments, poShowWithin, poResultCallback)
		  #EndIf
		  
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Function ComposeMessage(psTo As String, psBody As String, poAttachments() As FolderItem, poShowWithin As DesktopWindow, poResultCallback As ResultCallbackDelegate) As Boolean
		  #If TargetMacOS And Target64Bit Then
		    'This will perform the requested Sharing Service:
		    '- items from the parameters
		    '- poShowWithin: attach to that window (modally); nil -> sharing window will be independent
		    '- poResultCallback: reports back once Sharing Service has finished (success | error)
		    'Returns true - if NSSharingService could be invoked (not necessarily: has successfully shared the items!)
		    '               then wait for the callback with success or error
		    'Return false - if NSSharingService can't be invoked
		    Return PerformWithItems(NSSharingServiceName.ComposeMessage, psTo, "", psBody, poAttachments, poShowWithin, poResultCallback)
		  #EndIf
		  
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Private Sub Delegate_Implementation_DidFailToShare(id As Ptr, selectorPtr As Ptr, sharingServiceInstance As Ptr, items As Ptr, error As Ptr)
		  #Pragma unused id
		  #Pragma unused selectorPtr
		  #Pragma unused sharingServiceInstance
		  #Pragma unused items
		  
		  #If TargetMacOS And Target64Bit Then
		    //https://developer.apple.com/documentation/appkit/nssharingservicedelegate/1402710-sharingservice?language=objc
		    //- (void)sharingService:(NSSharingService *)sharingService didFailToShareItems:(NSArray *)items error:(NSError *)error;
		    
		    Declare Function code Lib "Foundation" Selector "code" ( ptrNSError As Ptr ) As Integer
		    'Declare Function domain Lib "Foundation" selector "domain" ( ptrNSError As Ptr ) As CFStringRef
		    Declare Function localizedDescription Lib "Foundation" Selector "localizedDescription" ( ptrNSError As Ptr ) As CFStringRef
		    
		    'get information out of the NSError
		    Var iError As Integer = code(error)
		    'Var sDomain As String = domain(error)
		    Var sError As String = localizedDescription(error)
		    
		    'invoke the callback with the result
		    If (mResultCallbackDelegate <> Nil) Then
		      mResultCallbackDelegate.Invoke(False, iError, sError)
		    End If
		    
		    mResultCallbackDelegate = Nil
		    mWeakRefShowWithinWindow = Nil
		  #EndIf
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Private Sub Delegate_Implementation_DidShareItems(id As Ptr, selectorPtr As Ptr, sharingServiceInstance As Ptr, items As Ptr)
		  #Pragma unused id
		  #Pragma unused selectorPtr
		  #Pragma unused sharingServiceInstance
		  #Pragma unused items
		  
		  #If TargetMacOS And Target64Bit Then
		    //https://developer.apple.com/documentation/appkit/nssharingservicedelegate/1402638-sharingservice?language=objc
		    //- (void)sharingService:(NSSharingService *)sharingService didShareItems:(NSArray *)items;
		    
		    'invoke the callback with the result
		    If (mResultCallbackDelegate <> Nil) Then
		      mResultCallbackDelegate.Invoke(True, 0, "")
		    End If
		    
		    mResultCallbackDelegate = Nil
		    mWeakRefShowWithinWindow = Nil
		  #EndIf
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Private Function Delegate_Implementation_SourceWindow(id As Ptr, selectorPtr As Ptr, sharingServiceInstance As Ptr, items As Ptr, sharingContentScope As Ptr) As Ptr
		  #Pragma unused id
		  #Pragma unused selectorPtr
		  #Pragma unused sharingServiceInstance
		  #Pragma unused items
		  
		  #If TargetMacOS And Target64Bit Then
		    //https://developer.apple.com/documentation/appkit/nssharingservicedelegate/1402679-sharingservice?language=objc
		    //- (NSWindow *)sharingService:(NSSharingService *)sharingService sourceWindowForShareItems:(NSArray *)items sharingContentScope:(NSSharingContentScope *)sharingContentScope;
		    
		    'configure that we're sharing item(s), and not partial or full content
		    Var mbSharingContentScope As MemoryBlock = sharingContentScope
		    mbSharingContentScope.Int32Value(0) = CType(NSSharingContentScope.Item, Int32)
		    
		    'return the Window, so that NSSharingService will show modally on it
		    If (mWeakRefShowWithinWindow <> Nil) And (mWeakRefShowWithinWindow.Value IsA DesktopWindow) Then
		      Return Ptr(DesktopWindow(mWeakRefShowWithinWindow.Value).Handle)
		    End If
		    
		    Return Nil
		  #EndIf
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Private Function NSSharingServiceName_ToString_ToString(SharingServiceName As NSSharingServiceName) As String
		  #If TargetMacOS And Target64Bit Then
		    //https://developer.apple.com/documentation/appkit/nssharingservicename?language=objc
		    
		    'return Xojo's Enum as a String value for NSSharingService
		    Select Case SharingServiceName
		    Case NSSharingServiceName.ComposeEmail
		      Return "com.apple.share.Mail.compose"
		    Case NSSharingServiceName.ComposeMessage
		      Return "com.apple.messages.ShareExtension"
		    Case NSSharingServiceName.SendViaAirDrop
		      Return "com.apple.share.AirDrop.send"
		    End Select
		  #EndIf
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Private Function PerformWithItems(psPerformWithNSSharingServiceName As NSSharingServiceName, psTo As String, psSubject As String, psBody As String, poAttachments() As FolderItem, poShowWithin As DesktopWindow, poResultCallback As ResultCallbackDelegate) As Boolean
		  mWeakRefShowWithinWindow = Nil
		  mResultCallbackDelegate = Nil
		  
		  #If TargetMacOS And Target64Bit Then
		    'store that for later (will be used by the Delegate implementations)
		    If (poShowWithin <> Nil) Then mWeakRefShowWithinWindow = New WeakRef(poShowWithin)
		    mResultCallbackDelegate = poResultCallback
		    
		    
		    'Separators supported: ; and ,
		    'Split Recipients to get an Array
		    Var sRecipients() As String = psTo.ReplaceAll(";", ",").Split(",")
		    For i As Integer = sRecipients.LastIndex DownTo 0
		      sRecipients(i) = sRecipients(i).Trim
		    Next
		    
		    Var sSubject As String = psSubject.Trim
		    Var sBody As String = psBody.Trim
		    
		    Var oAttachments() As FolderItem = poAttachments
		    
		    //https://developer.apple.com/documentation/appkit/nssharingservice?language=objc
		    //Declares: NSSharingService
		    Var sPerformWithNSSharingServiceName As String= NSSharingServiceName_ToString_ToString(psPerformWithNSSharingServiceName)
		    
		    Declare Function NSClassFromString Lib "Foundation" (className As CFStringRef) As Ptr
		    Declare Function sharingServiceNamed Lib "AppKit" Selector "sharingServiceNamed:" (NSSharingServiceClass As Ptr, serviceName As CFStringRef) As Ptr
		    
		    Declare Sub setRecipients Lib "AppKit" Selector "setRecipients:" (NSSharingServiceInstance As Ptr, obj As Ptr)
		    Declare Sub setSubject Lib "AppKit" Selector "setSubject:" (NSSharingServiceInstance As Ptr, subject As CFStringRef)
		    Declare Function canPerformWithItems Lib "AppKit" Selector "canPerformWithItems:" (NSSharingServiceInstance As Ptr, obj As Ptr) As Boolean
		    Declare Sub performWithItems Lib "AppKit" Selector "performWithItems:" (NSSharingServiceInstance As Ptr, obj As Ptr)
		    Declare Sub setDelegate Lib "AppKit" Selector "setDelegate:" (id As Ptr, ptrToDelegate As Ptr)
		    
		    'NSSharingService instance
		    Var ptrNSSharingServiceClass As Ptr = NSClassFromString("NSSharingService")
		    Var ptrNSSharingServiceInstance As Ptr = sharingServiceNamed(ptrNSSharingServiceClass, sPerformWithNSSharingServiceName)
		    
		    
		    //Declares: NSMutableArray
		    Declare Function alloc Lib "Foundation" Selector "alloc" (NSClass As Ptr) As Ptr
		    Declare Function init Lib "Foundation" Selector "init" (NSClass As Ptr) As Ptr
		    Declare Sub addObject_String Lib "Foundation" Selector "addObject:"(NSMutableArrayClass As Ptr, anObject As CFStringRef)
		    Declare Sub addObject_Ptr Lib "Foundation" Selector "addObject:"(NSMutableArrayClass As Ptr, anObject As Ptr)
		    
		    Var ptrMutableArrayClass As Ptr = NSClassFromString("NSMutableArray")
		    
		    'Build Recipients Array
		    Var ptrRecipients As Ptr = init(alloc(ptrMutableArrayClass))
		    For Each sRecipient As String In sRecipients
		      addObject_String(ptrRecipients, sRecipient)
		    Next
		    
		    'Build Items Array for Content (Body and Attachments)
		    Var ptrItems As Ptr = init(alloc(ptrMutableArrayClass))
		    addObject_String(ptrItems, sBody)
		    
		    'add Attachments to Items Array
		    Declare Function fileURLWithPath Lib "Foundation" Selector "fileURLWithPath:" ( ptrNSURLClass As Ptr, path As CFStringRef ) As Ptr
		    For Each oAttachFolderItem As FolderItem In oAttachments
		      'just existing Files, no Folders
		      If (oAttachFolderItem = Nil) Or (oAttachFolderItem.Exists = False) Or oAttachFolderItem.IsFolder Then Continue
		      
		      'NSURL for Attachment
		      Var ptrNSURLClass As Ptr = NSClassFromString("NSURL")
		      Var ptrAttachment As Ptr = fileURLWithPath(ptrNSURLClass, oAttachFolderItem.NativePath)
		      
		      addObject_Ptr(ptrItems, ptrAttachment)
		    Next
		    
		    'NSSharingService: set Subject, Recipient
		    setSubject(ptrNSSharingServiceInstance, sSubject)
		    setRecipients(ptrNSSharingServiceInstance, ptrRecipients)
		    
		    'setup NSSharingServiceDelegate
		    If (ptrDelegateClass = Nil) Then ptrDelegateClass = BuildNSSharingServiceDelegateClass
		    
		    'attach delegate
		    setDelegate(ptrNSSharingServiceInstance, ptrDelegateClass)
		    
		    //https://developer.apple.com/documentation/appkit/nssharingservice/1402662-canperformwithitems?language=objc
		    //https://developer.apple.com/documentation/appkit/nssharingservice/1402669-performwithitems?language=objc
		    'NSSharingService: Perform with Content-Items
		    If canPerformWithItems(ptrNSSharingServiceInstance, ptrItems) Then
		      performWithItems(ptrNSSharingServiceInstance, ptrItems)
		      Return True
		    Else
		      Var sError As String = "NSSharingService can't perform '" + sPerformWithNSSharingServiceName + "' with the assigned items."
		      System.DebugLog sError
		      Break
		      
		      Return False
		    End If
		    
		  #EndIf
		  
		  Return False
		End Function
	#tag EndMethod

	#tag DelegateDeclaration, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Private Delegate Sub ResultCallbackDelegate(pbSuccess As Boolean, piErrorCode As Integer, psErrorMessage As String)
	#tag EndDelegateDeclaration

	#tag Method, Flags = &h0, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Function SendViaAirDrop(poAttachments() As FolderItem, poShowWithin As DesktopWindow, poResultCallback As ResultCallbackDelegate) As Boolean
		  #If TargetMacOS And Target64Bit Then
		    'This will perform the requested Sharing Service:
		    '- items from the parameters
		    '- poShowWithin: attach to that window (modally); nil -> sharing window will be independent
		    '- poResultCallback: reports back once Sharing Service has finished (success | error)
		    'Returns true - if NSSharingService could be invoked (not necessarily: has successfully shared the items!)
		    '               then wait for the callback with success or error
		    'Return false - if NSSharingService can't be invoked
		    Return PerformWithItems(NSSharingServiceName.SendViaAirDrop, "", "", "", poAttachments, poShowWithin, poResultCallback)
		  #EndIf
		  
		  Return False
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Private mResultCallbackDelegate As ResultCallbackDelegate
	#tag EndProperty

	#tag Property, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Private mWeakRefShowWithinWindow As WeakRef
	#tag EndProperty

	#tag Property, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Private ptrDelegateClass As Ptr
	#tag EndProperty


	#tag Enum, Name = NSSharingContentScope, Type = Int32, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		Item=0
		  Partial=1
		Full=2
	#tag EndEnum

	#tag Enum, Name = NSSharingServiceName, Type = Integer, Flags = &h21, CompatibilityFlags = API2Only and ( (TargetDesktop and (Target32Bit or Target64Bit)) )
		ComposeEmail=1
		  ComposeMessage=2
		SendViaAirDrop=3
	#tag EndEnum


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
