<%
'**********************************************
'**********************************************
'               _ _                 
'      /\      | (_)                
'     /  \   __| |_  __ _ _ __  ___ 
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/ 
' 
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
' LAST UPDATE: 28.10.2022 15:33 @badursun
'**********************************************

Class Facebook_oAuth_Plugin
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT, PLUGIN_ICON, PLUGIN_REMOVABLE, PLUGIN_ROOT, PLUGIN_FOLDER_NAME, PLUGIN_AUTOLOAD

	Private FB_AUTH_URL, FB_TOKEN_URL, FB_DATA_URL
	Private FB_APP_ID, FB_APP_SECRET, FB_CALLBACK_URL, FB_SCOPE
	Private FB_REG_CODE
	Private GET_RESPONSE
	Private FB_TOKEN, FB_FIELDS

	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"
		
		' Check Register
		'------------------------------
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If

		' Check And Create Table
		'------------------------------
		' Dim PluginTableName
		' 	PluginTableName = "tbl_plugin_" & PLUGIN_DB_NAME
    	
  '   	If TableExist(PluginTableName) = False Then
		' 	DebugTimer ""& PLUGIN_CODE &" table creating"
    		
  '   		Conn.Execute("SET NAMES utf8mb4;") 
  '   		Conn.Execute("SET FOREIGN_KEY_CHECKS = 0;") 
    		
  '   		Conn.Execute("DROP TABLE IF EXISTS `"& PluginTableName &"`")

  '   		q="CREATE TABLE `"& PluginTableName &"` ( "
  '   		q=q+"  `ID` int(11) NOT NULL AUTO_INCREMENT, "
  '   		q=q+"  `FILENAME` varchar(255) DEFAULT NULL, "
  '   		q=q+"  `FULL_PATH` varchar(255) DEFAULT NULL, "
  '   		q=q+"  `COMPRESS_DATE` datetime DEFAULT NULL, "
  '   		q=q+"  `COMPRESS_RATIO` double(255,0) DEFAULT NULL, "
  '   		q=q+"  `ORIGINAL_FILE_SIZE` bigint(20) DEFAULT 0, "
  '   		q=q+"  `COMPRESSED_FILE_SIZE` bigint(20) DEFAULT 0, "
  '   		q=q+"  `EARNED_SIZE` bigint(20) DEFAULT 0, "
  '   		q=q+"  `ORIGINAL_PROTECTED` int(1) DEFAULT 0, "
  '   		q=q+"  PRIMARY KEY (`ID`), "
  '   		q=q+"  KEY `IND1` (`FILENAME`) "
  '   		q=q+") ENGINE=MyISAM DEFAULT CHARSET=utf8; "
		' 	Conn.Execute(q) : q = ""

  '   		Conn.Execute("SET FOREIGN_KEY_CHECKS = 1;") 

		' 	' Create Log
		' 	'------------------------------
  '   		Call PanelLog(""& PLUGIN_CODE &" i??in "& PluginTableName &" tablosu olu??turuldu", 0, ""& PLUGIN_CODE &"", 0)
			
		' 	DebugTimer ""& PLUGIN_CODE &" "& PluginTableName &" table created"
  '   	End If

		' Register Settings
		'------------------------------
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE)
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "Facebook_oAuth_Plugin")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "5")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", PLUGIN_FOLDER_NAME)

		a=GetSettings(""&PLUGIN_CODE&"_ACTIVE", "0")
		a=GetSettings(""&PLUGIN_CODE&"_APP_ID", "")
		a=GetSettings(""&PLUGIN_CODE&"_APP_SECRET", "5")

		' Register Settings
		'------------------------------
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public sub LoadPanel()
		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "SHOW:HowTo" Then
			Call PluginPage("Header")

			With Response 
				.Write "<p>"
				.Write "	<ol>"
				.Write "		<li>Bir Facebook kullan??c?? hesab??na ihtiyac??n??z olacak</li>"
				.Write "		<li>https://developers.facebook.com/ adresine gidin.</li>"
				.Write "		<li>Get Started veya Facebook Login butonuna t??klay??n.</li>"
				.Write "		<li>Hesab??n??z?? kaydedin ve do??rulay??n</li>"
				.Write "		<li>Facebook uygulaman??z i??in bir ad girin (???RabbitCMS Kurumsal??? gibi bir ad se??ebilirsiniz). Bu ad, Facebook Connect ??zerinden uygulamada bir hesap olu??turmak istediklerinde kullan??c??lara g??sterilecek.</li>"
				.Write "		<li>???Bir Senaryo se?????, (Select a Scenario) sayfas?? alt??nda ???Facebook Giri????? (Integrate Facebook Login) se??in</li>"
				.Write "		<li>Soldaki men??de Ayarlar (Settings)> Temel???e (Basic) gidin .</li>"
				.Write "		<li>Kullanaca????n??z Uygulama Kimli??i (App ID) ve ??ifresi (App Secret).</li>"
				.Write "		<li>E??er ekli de??ilse sol men??de Products alt??nda Facebook Login eklenmeli.</li>"
				.Write "		<li>??lgili tan??mlamalar??, domain adresinizi ekleyin.</li>"
				.Write "		<li>T??m detaylar tamamland?? ise ??st b??l??mde App ID yan??nda bulunan anahtar?? Live hale getirin.</li>"
				.Write "	</ol>"
				.Write "</p>"
			End With

			Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("input", ""& PLUGIN_CODE &"_APP_ID", "App ID", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("input", ""& PLUGIN_CODE &"_APP_SECRET", "App Secret", "", TO_DB)
			.Write "    </div>"
			.Write "</div>"
			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=SHOW:HowTo"" class=""btn btn-sm btn-primary"">"
			.Write "        	Nas??l App Olu??turulur?"
			.Write "        </a>"
			If Len(FB_APP_ID) > 2 Then
			.Write "        <a target=""_blank"" href=""https://developers.facebook.com/apps/"& FB_APP_ID &"/"" class=""btn btn-sm btn-info"">"
			.Write "        	Facebook Developers da Uygulamay?? A??"
			.Write "        </a>"
			End If
			.Write "    </div>"
			.Write "</div>"
		End With
	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	private sub class_initialize()
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
    	PLUGIN_NAME 			= "Facebook oAuth Plugin"
    	PLUGIN_CODE  			= "FACEBOOK_OAUTH"
    	PLUGIN_DB_NAME 			= "fboauth_log" ' tbl_plugin_XXXXXXX
    	PLUGIN_VERSION 			= "1.1.4"
    	PLUGIN_CREDITS 			= "@badursun Anthony Burak DURSUN"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/Facebook-oAuth-Plugin"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_FOLDER_NAME 		= "Facebook-oAuth-Plugin"
    	PLUGIN_ICON 			= "zmdi-facebook"
    	PLUGIN_REMOVABLE 		= True
    	PLUGIN_AUTOLOAD 		= False
    	PLUGIN_ROOT 			= PLUGIN_DIST_FOLDER_PATH(This)
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
		
		FB_APP_ID 			= GetSettings(""& PLUGIN_CODE &"_APP_ID", "")
		FB_APP_SECRET 		= GetSettings(""& PLUGIN_CODE &"_APP_SECRET", "")
		FB_AUTH_URL 		= "https://www.facebook.com/dialog/oauth"
		FB_TOKEN_URL 		= "https://graph.facebook.com/oauth/access_token"
		FB_CALLBACK_URL 	= DOMAIN_URL & "/oauth/facebook/"
		FB_DATA_URL 		= "https://graph.facebook.com/me?access_token="
		FB_SCOPE 			= "public_profile,email"
		FB_FIELDS 			= "id,first_name,middle_name,last_name,email,picture"
		FB_REG_CODE 		= Null
		FB_TOKEN 			= Null

    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Register App
    	'-------------------------------------------------------------------------------------
    	class_register()

    	'-------------------------------------------------------------------------------------
    	' Hook Auto Load Plugin
    	'-------------------------------------------------------------------------------------
    	If PLUGIN_AUTOLOAD_AT("WEB") = True Then 

    	End If
	end sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private sub class_terminate()

	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------
	Public Property Get PluginCode() 		: PluginCode = PLUGIN_CODE 					: End Property
	Public Property Get PluginName() 		: PluginName = PLUGIN_NAME 					: End Property
	Public Property Get PluginVersion() 	: PluginVersion = PLUGIN_VERSION 			: End Property
	Public Property Get PluginGit() 		: PluginGit = PLUGIN_GIT 					: End Property
	Public Property Get PluginDevURL() 		: PluginDevURL = PLUGIN_DEV_URL 			: End Property
	Public Property Get PluginFolder() 		: PluginFolder = PLUGIN_FILES_ROOT 			: End Property
	Public Property Get PluginIcon() 		: PluginIcon = PLUGIN_ICON 					: End Property
	Public Property Get PluginRemovable() 	: PluginRemovable = PLUGIN_REMOVABLE 		: End Property
	Public Property Get PluginCredits() 	: PluginCredits = PLUGIN_CREDITS 			: End Property
	Public Property Get PluginRoot() 		: PluginRoot = PLUGIN_ROOT 					: End Property
	Public Property Get PluginFolderName() 	: PluginFolderName = PLUGIN_FOLDER_NAME 	: End Property
	Public Property Get PluginDBTable() 	: PluginDBTable = IIf(Len(PLUGIN_DB_NAME)>2, "tbl_plugin_"&PLUGIN_DB_NAME, "") 	: End Property
	Public Property Get PluginAutoload() 	: PluginAutoload = PLUGIN_AUTOLOAD 			: End Property

	Private Property Get This()
		This = Array(PluginCode, PluginName, PluginVersion, PluginGit, PluginDevURL, PluginFolder, PluginIcon, PluginRemovable, PluginCredits, PluginRoot, PluginFolderName, PluginDBTable, PluginAutoload)
	End Property
	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public Property Get GetToken()
		GetToken = FB_AUTH_URL & "?response_type=code&client_id="& FB_APP_ID &"&redirect_uri="& Server.URLEncode(FB_CALLBACK_URL) &"&scope="& FB_SCOPE &"&state="& ConvertToUnixTimeStamp(Now())
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Let SetCode(TheCode)
		FB_REG_CODE = TheCode
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Let SetToken(TheCode)
		FB_TOKEN = TheCode
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Get access_token()
		If Len(FB_REG_CODE) < 10 Then 
			access_token = Array(304, "The Code Not Exist")
			Exit Property
		End If
		
		Dim GET_TOKEN_URL
			GET_TOKEN_URL = FB_TOKEN_URL &"?client_id="& FB_APP_ID &"&redirect_uri="& Server.URLEncode(FB_CALLBACK_URL) &"&client_secret="& FB_APP_SECRET &"&code="& FB_REG_CODE
		
		Dim tmp_response
			tmp_response = XMLHttp(GET_TOKEN_URL, "GET", "")

		' Yan??t?? par??ala ve ayr????t??r
		'----------------------------------------------
		Set parseJsonData = New aspJSON
			parseJsonData.loadJSON( tmp_response(1) )
		
			If IsNull(parseJsonData.data("access_token")) = True Then 
				access_token = Array(300, parseJsonData.data("error").item("message") )
			Else 
				access_token = Array(200, parseJsonData.data("access_token") )
			End If
		Set parseJsonData = Nothing
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Get MeData()
		If Len(FB_TOKEN) < 10 Then 
			MeData = Array(304, "Token Not Exist")
			Exit Property
		End If
		
		Dim GET_DATA_URL
			GET_DATA_URL = FB_DATA_URL & FB_TOKEN & "&fields="& FB_FIELDS
		
		Dim tmp_response
			tmp_response = XMLHttp(GET_DATA_URL, "GET", "")

		' Yan??t?? par??ala ve ayr????t??r
		'----------------------------------------------
		Set parseJsonData = New aspJSON
			parseJsonData.loadJSON( tmp_response(1) )
		
			If TypeName(parseJsonData.data("error")) = "Empty" Then 
				MeData = Array(200, tmp_response(1) )
			Else 
				MeData = Array(300, parseJsonData.data("error").item("message") )
			End If
		Set parseJsonData = Nothing
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Property Get XMLHttp(Uri, xType, Data)
		On Error Resume Next

		' Send Data
		'------------------------------------------------
	    Set objXMLhttp = Server.Createobject("MSXML2.ServerXMLHTTP")
            objXMLhttp.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
            objXMLhttp.setTimeouts 5000, 5000, 10000, 10000 'ms
            ' objXMLhttp.setRequestHeader "X-Cms-UniqueId", SETTINGS_CMS_UNIQUE_ID
			objXMLhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
			objXMLhttp.open xType, Uri, false
			objXMLhttp.send Data
			
			GET_RESPONSE = objXMLhttp.responseText
			XMLHttp = Array(objXMLhttp.Status, objXMLhttp.responseText)
		
		CreateLog "facebook.oAuth.XMLHttp", JSONTurkish(Data), JSONTurkish( objXMLhttp.responseText ), objXMLhttp.Status, UCase(xType)
	    
	    Set objXMLhttp = Nothing
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
End Class 
%>
