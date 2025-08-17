

On Error Resume Next

'****Defini��o de vari�veis
Dim aHeader(10)
Dim aInterfacesSubKeys
Dim aIPAddress
Dim bConectadoSequoia
Dim bContinuaTentando
Dim bEncontrado
Dim bInconsistencia2
Dim bNatura
Dim bRede
Dim bRedeDesconhecida
Dim cBIOS
Dim cContasLocais
Dim cDisks
Dim cIPConfigSet
Dim cPings
Dim iCodigoMaquinaNova
Dim iLineLog
Dim iTentativa
Dim iTentativasConexao
Dim oBIOS
Dim oConnectionSequoia
Dim oDictionary
Dim oDisk
Dim oFile
Dim oFileLog
Dim oFSO
Dim oIPConfig
Dim oNetwork
Dim oPing
Dim oRecordSetAdministradores
Dim oRecordsetBackupDados
Dim oRecordSetDomainControllers
Dim oRecordSetMaquinas
Dim oRecordSetRangeIP
Dim oRecordSetSiteDomainControllers
Dim oReg
Dim oShell
Dim oUsuarioLocal
Dim oWMI
Dim sArquivoLog
Dim sCaminhoBackupDados
Dim sCaminhoLog
Dim sDatabaseServer
Dim sDhcpDomain
Dim sDhcpIPAddress
Dim sDominio
Dim sDrive
Dim sProximaLetraDisponivel
Dim sSenhaAreaBackup
Dim sSenhaRede
Dim sScriptPath
Dim sSerialNumber
Dim sUsuarioAreaBackup
Dim sUsuarioRede

'****Defini��o de constantes que auxiliam a cria��o do Log
Const SUCCESS = 0
Const ERRORL = 1
Const WARNING = 2
Const INFORMATION = 3

'****Defini��o de constantes que auxiliam manipula��o do Registro do Windows
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

'****Inicia objeto oFSO, oWMI, oShell, oReg e oNetwork
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set oShell = WScript.CreateObject("WScript.Shell")
Set oReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
Set oNetwork = CreateObject("Wscript.Network")

'****Cabe�alho do Log
aHeader(0) = "---------------------------------------------------------"
aHeader(1) = "        EIW - Engenharia de Infraestrutura Wintel"
aHeader(2) = "                     Projeto Sequ�ia"
aHeader(3) = ""
aHeader(4) = " Script:	Captura de Estado da m�quina antiga de um"
aHeader(5) = "		usu�rio em migra��o"
aHeader(6) = ""
aHeader(7) = " Data:		16/07/2010"
aHeader(8) = " Vers�o:	1.0"
aHeader(9) = "---------------------------------------------------------"

'****Define nome do arquivo de log
sArquivoLog = "LogCapturaEstado.log"

'****Define o caminho para salvar o arquivo de log (HomeDrive para Windows 7 ou superiores e Temp para anteriores ao Windows 7)
sCaminhoLog = ""
Set cSistemaOperacional = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each oSistemaOperacional in cSistemaOperacional
	sSOVersao = oSistemaOperacional.Version
Next
If CInt(Mid(sSOVersao, 1, 3)) >= 61 Then 'Windows 7
	sCaminhoLog = oShell.ExpandEnvironmentStrings("%HOMEDRIVE%")
Else 'Anteriores ao Windows 7
	sCaminhoLog = oShell.ExpandEnvironmentStrings("%TEMP%")
	If sCaminhoLog = "" Then
		sCaminhoLog = oShell.ExpandEnvironmentStrings("%TMP%")
	End If
End If

'****Verifica se o arquivo de log j� existe
'****Se n�o existir, ent�o cria, abre para adi��o de novas linhas e configura o arquivo de log
'****Sen�o abre o arquivo de log para adi��o de novas linhas
If Not oFSO.FileExists(sCaminhoLog & "\" & sArquivoLog) Then

	'****Se o arquivo n�o existe ent�o ele � criado e aberto para a adi��o de novas linhas e o cabe�alho do log � inserido
	oFSO.CreateTextFile(sCaminhoLog & "\" & sArquivoLog)
	Set oFileLog = oFSO.OpenTextFile(sCaminhoLog & "\" & sArquivoLog, 8, True)
	For i = 0 To UBound(aHeader)
		s_Log INFORMATION, "[HEADER] " & aHeader(i)
	Next
	s_Log SUCCESS, "[Main Code] Arquivo de log criado - " & sCaminhoLog & "\" & sArquivoLog

	'****Oculta o arquivo
	Set oFile = oFSO.GetFile(sCaminhoLog & "\" & sArquivoLog)
	oFile.Attributes = 2
	s_Log SUCCESS, "[Main Code] Atributo ""Oculto"" do arquivo de log ativado - " & sCaminhoLog & "\" & sArquivoLog
Else

	'****Se o arquivo existe ent�o ele � aberto para a adi��o de novas linhas e o cabe�alho do log � inserido
	Set oFileLog = oFSO.OpenTextFile(sCaminhoLog & "\" & sArquivoLog, 8, True)
	For i = 0 To UBound(aHeader)
		s_Log INFORMATION, "[HEADER] " & aHeader(i)
	Next
	s_Log INFORMATION, "[Main Code] Arquivo de log localizado - " & sCaminhoLog & "\" & sArquivoLog
End If

'****Inicia vari�vel sScriptPath
sScriptPath = Replace(WScript.ScriptFullName, "\" & WScript.ScriptName, "")

'****Inicia vari�vel sCaminhoBackupDados
sCaminhoBackupDados = ""

'****Inicia vari�veis sDominio, sProximaLetraDisponivel, sSenhaAreaBackup, sUsuarioAreaBackup, sSenhaRede e sUsuarioRede
sDominio = ""
sProximaLetraDisponivel = ""
sSenhaAreaBackup = ""
sSenhaRede = ""
sUsuarioAreaBackup = ""
sUsuarioRede = ""

'****Define servidor do banco de dados Sequ�ia de acordo com IP (Rede Natura ou Rede laborat�rio Sequoia)
aIPAddress = Split("0.0.0.0",".")
bRede = True
bRedeDesconhecida = False
'Set cIPConfigSet = oWMI.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
'For Each oIPConfig in cIPConfigSet
'	aIPAddress = Split(oIPConfig.IPAddress(0), ".")
'Next
oReg.EnumKey HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces", aInterfacesSubKeys
For i = 0 To UBound(aInterfacesSubKeys)
	sDhcpDomain = ""
	sDomain = ""
	sDhcpNameServer = ""
	oReg.GetStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & aInterfacesSubKeys(i), "DhcpDomain", sDhcpDomain
	oReg.GetStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & aInterfacesSubKeys(i), "Domain", sDomain
	oReg.GetStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & aInterfacesSubKeys(i), "DhcpNameServer", sDhcpNameServer
	If (sDhcpDomain = "br.natura" or sDomain = "natura.com.br") and sDhcpNameServer <> "" Then
		oReg.GetStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\" & aInterfacesSubKeys(i), "DhcpIPAddress", sDhcpIPAddress
		aIPAddress = Split(sDhcpIPAddress, ".")
		Exit For
	End If
Next
'If aIPAddress(0) = "172" Then
'	sDatabaseServer = "AS3K98BR"
'ElseIf aIPAddress(0) & aIPAddress(1) & aIPAddress(2) = "19216895" Then
'	sDatabaseServer = "WS2008SRV1"
'ElseIf aIPAddress(0) & aIPAddress(1) = "00" Then
'	bRede = False
'Else
'	bRedeDesconhecida = True
'End If
If aIPAddress(0) & aIPAddress(1) & aIPAddress(2) = "19216895" Then
	sDatabaseServer = "WS2008SRV1"
ElseIf aIPAddress(0) & aIPAddress(1) = "00" Then
	bRede = False
Else
	sDatabaseServer = "AS3K98BR"
End If


If bRede and Not bRedeDesconhecida Then

	'****Verifica se consegue "pingar" o servidor do banco de dados Sequ�ia e escreve o resultado no Log
	s_Log INFORMATION, "[Main Code] Teste de Ping para o servidor do banco de dados Sequ�ia: " & Replace(sDatabaseServer, "\SQLEXPRESS", "") & "."
	bNatura = False
	Set cPings = oWMI.ExecQuery ("Select * from Win32_PingStatus Where Address = '" & Replace(sDatabaseServer, "\SQLEXPRESS", "") & "'")
	For Each oPing in cPings
		If oPing.StatusCode = 0 Then
			s_Log SUCCESS, "[Main Code] Ping OK:			" & Replace(sDatabaseServer, "\SQLEXPRESS", "")
			bNatura = True
		Else
			s_Log ERRORL, "[Main Code] Ping Falhou:		" & Replace(sDatabaseServer, "\SQLEXPRESS", "")
		End If
	Next

	'****Se conseguir resposta do servidor do banco de dados Sequ�ia, ent�o est� na rede Natura e pode continuar a reunir informa��es sobre a m�quina
	If bNatura Then
		s_Log INFORMATION, "[Main Code] Conectado na rede da Natura."

		'****Identifica o n�mero de s�rie do equipamento
		Set cBIOS = oWMI.ExecQuery("SELECT * FROM Win32_BIOS")
		For Each oBIOS in cBIOS
			sSerialNumber = oBIOS.SerialNumber
		Next

		'****Inicia a conex�o com o banco de dados Sequ�ia (5 tentativas caso ocorra erro)
		s_Log INFORMATION, "[Main Code] Iniciando conex�o com banco de dados Sequ�ia..."
		iTentativa = 1
		iTentativasConexao = 5
		bContinuaTentando = True
		bConectadoSequoia = False
		Err.Clear
		Set oConnectionSequoia = CreateObject("ADODB.Connection")
		While bContinuaTentando and iTentativa <= iTentativasConexao
			oConnectionSequoia.Open "Provider=SQLOLEDB.1;Data Source=" & sDatabaseServer & ";Trusted_Connection=No;Initial Catalog=DB_SEQUOIA;User ID=usr_sequoia;Password=P@B	CtJzU2qCC|@qUO
UsvXxTpzV(NC#bP*{n 09ohe=
	U	vs`,Dj9~SqE~~ [u MaiW Cf#P% f{Kc]:#79:fB}_nlf &6{&adT S2[C�i_ Sa[%K~r5%A|a9j2[~�Pero[" dBO#eS%c9i	:B=9L-'
	vs		`w7g 3cFO;h}T/Oy,[L vcs=9oJ60[UKrviDorAv^U"Bd9LD_%*}cC2VP*vK]
UL^	sKwogT3yFQRZA#1fM'{L%ZaAJTu7#e]  an_D,DeB3c#o|Av^~2uC7)a'
^UL^bCTJec%c#oneQuP|*,_{PGCe
sU^vbCoWti=CH#eS%cSdPTJ,p:kLP
vs	kcP
;L		|mKn]a9i	:Bn9sPWStc]|	HB+ 1
	sv^/fB|mKn]:9O^:T>,)TK=tH]s\asq4neOcD #hKn
v^;LsiaLTXTw%[ORLJ "BZH)nB!T#e%T+VGfT#P[CD=e#�P{$om9han_TBd6 #aDfL,~6u0�"a`{p*VD n�m6roT',g +]G-NG#1PG
v^U	+=d,/k
	;LEn}9bf
v	12S&

v	u!YaG/ocaciz:B7[eQ"A3al2S-TT=D,IaS_o,D6,daJ4s ^WQu�)aBa]]*\�i{}TBsW'TqPero[deTVP*iKT#*r_T|H$|}*V[sKTh�,xt�.caJ5st]TBdP Pe|#D
LsX+nN4=]V*do 7 F:m	6
s;bf[h7{2_9HDoV2qqPsH Tb6n
;v	wrz.ukKHG
;v	~6%[7[eco*dS29v_qC|{*s[&BsG2:9PQb^2c-'/:DO7E.R2NDrDsKtiV
;LsT[eN4]DUKtMaBui=*	(O32{B"np/kqPTr,\RfN &I42Ma['in:iBW?E[E[^KVA5k?CmX6]["bKE "" dB	nez|cmNG#1PGTdBu""wTosPJ~ecrson^WQuPi*,[.w,E
;v	TY297zdSe]MauCOWaL-hDv6l"Vi%
^U	b8 k*t2NuFherT_B0[TheW
^;LsIhi$6T79 oR6co]#U6tZ:[CiW:L2BilBHWdB9o-[hkncTJtr:JD
	v	Usbo94;WloGj^6-Zaqu)naC-g)em}i."n2zOck9CwIezL)2S5Fue9<> C~Kr)amNG#1PG{PbKn
sU^v		ozecfz&ne9NcQu)=*	4Nf|Pep%
UL^		BGse
v	U	v	|e*-Gs#TSiT{&[7[eco*dS29v_qC|{*s(l"P$}C.uma9]iwPJOo"58Vak0K
	v	Usv}BJ_TStG5}P,J TrGe
v^U	vp{# )8
;ssv{6n#
^ULknd9Wf
;vEWdBIO

Lsur*H!^6,KncoWtr:z,P Ku0"p_#K~rfTS7[b*=c7[jP dcjosT~KqG�i*,[2S-�T{}W4i{6To7PiniP, "Lq�*iDTWBs6=hH9\:zH[RC= :|{&o ~ari\rL
	v'aYrG~6T{�D WJ_P~9rar[e CK,6s9|^Kr[_D~W_%*&P lfm,P{}anN4 d29#aDoL n2Qq�A5 9Knr�4T6#"be lenC*j6mB|{4o*#*~JfTQq6 �B=e$6	�riT{ex|i9i* laD:L-G4TJD BC")E*men]o \*V_ KOWlu]:z,cT_*E]uz: &6{Pstcjo
;vIO 1EW_D~rt:JD mM2W

	U	'YrGaEp2NCt_T1FT_fBHIa"Oo,|6,o WCui\cPeWtD Of",WJ_TStG5}P,So b_ncfB&6 #:JDs[^KY0�|:
U	vC_"PX,SUq0ES^'B"RM*iWT7J609+q0s\_wKnto[lo_*F)z*}TBn_T9HX2k*,ZaQ"i~_2"
L		CK/o0 bN\i[v/o1Zy,9/BZH"n CPde0BM�lezf9#e[C�zOWAsvu[&BCSP*sHlN0Tbe]

'v	UJrGH!3W4i{6To7PiniP, "Lq�*iDTWBs6=hH9}2BV6dKT(&P{HdmAJis%G*dPrBdPTKY0s\cPe{]f4,3ara[Ru=B:| #f9Vc*|3-i T1q|c*=d7[s~foGTa��oTJD WaBt_hKFc{)JPi{sC]V*dor6s {{Oxo}}F"n)C9Vc}fz,7 )V
"L^	EGt.CkW*r
(	UsVPr{f!KcTt}nP9Adm)niC9V_dD]WL 7TVW:%KxIjK_t4iiJOD,8Re_Tzd|e9"4
6;LsT[eN4]DUKtAdlin|L-*a#fGKs(i3P{TLVksEP G[m%OM9jbo-/#m)n"s]]*&Tt2iBW+3;w,{ipoydm|SO|tz:JDr[&B5O TDsPnS2c-)4~Se['oi:'B2J $
Jv^;4;WloGj^6-}dmiWis%zHDoz2i-MPMKgA]C9
'vs	XO{krr4cumhWz 7 g #ZK~
J;v	;S|q*riozed2Bn[o[2NDrD^K-/}#"~)s9]a&PtPs.psel}i."so&iW)#wAJ|i9rcjf*ui.Vacue
6U	vsiVeWZ*%W}2Bn[o[2c7*jUet/jmi=ALt*a#o*2L2ps2$#s6/^6~haAdlin|L-*a#fG')(r*F02
6U	vsscs4j S(0CE^~w i[Za)=BsTj2tBCTJ%_,#e dPm�n|D,xaz:9Kx6_C��79}fL,9oP:n&P,inr6rnfiBd6s9e[ClVAF%9QuW{}6EKndel d2BHGtK=r"c_��fB~TT}Dw�WiDTe,6GPva��T{deT#zi	im�g)fL,Aj2{9ixs__&* col s"lP|sD-O
"
6;ssv6a*rYDPOs~e J4m�n|TBd6 *u]2S-Aa:��TBdc{_P~9a bGsc:S&P DTJDm�W|D,N:}*	]r*}o,x5Va T{DCTJD |i9e[}D,WC"A3aF6=]7
'	U		pzV(Cm2cz
Jv^;s^K-[o[2c7*jUet~ste3TPa)noW%z7$G2GL _{!*P*teOIje_94iAdi7q.z2l7G}CK-i)
'^UL^oRWaor}~Ktni9eofPHAJ!TStG4kcPzs.OxenT'UwL+!mB*[l[xhT}17(S"%eJPTHinq4nt]Tml6rL 1a+%B{!T#i]4^)-K = i &TD%6cD]JVe]N*Y0|=*	(F"2l&|UuCoJsgo^A9ei)-V_kCP'{fqDn{6_]ODnSeBuo|*L[2wTE
"sv^;f;K$Pr#^e-ns-eDTTai=qDn]rDlc2z	4yf^KFAtC]
'		U	I8Bk*r-90Pb6]Bn9GT{b6n
'^UL^	EGt.CkW*r
(	Usv^~6%9DRWaf*&VetDPma|SsPn9]Tml6]L,_T!zP_tKib86a-("/$ODw4[e9ozd|29u5
uv	;LsP%KcorDSe%d7la"=qDn]]DF$2]L2QpK= un3"ECm{* l!fM[d1o(3Dwcs=qDnrtfcFKrs 1HE;+,uo#|]DDuTJ,OTdB7zelfr&n6-Sir6Do#c"nuoSt*fmFWtC4XiWG}|4'Cod)go3u4.`:$CeJTDsT==K$]iD=SPB'7ia'{2,TE
"	v	UsD%WafG#SW]3Pw*inCPnt]DFcezC4Zo	2XOGC%
"	vs	^/k,ErG8Nu#XKr[=B0[PhP{
uv	;LsU	domiWioTJ,PRK_Tzdn29JT#:"~uoS%r7cGPrs4miekJL(iDDm)="7O`-U*l06
6v			UElCK
'vs;v	UCa"TFT+%zO[,,,iqvai{{Co}Wt �oBfP|BETC�^Kl9s}6~9ifi9arTD,DoP�=AD D2BH0%2S-)c*��f &_{$onr5 p:G* 6xKcG��fB&TTNDmcJ}P	Bint6rnfL,DeL%WBs9]"ErTuCP[dK\e~D6w dW{au%WSt)c*��o[=D,J4#�{"o96T6FKva��o[deT3V)v"k�]"o|-'
Jsv^U	+=d,/k
';L		sT[e9ozdn29JTT:ASCTJ%*7mler|.CkD	6
J;v	Upm	W
(^U	vsscs4j E!YOR,'B"RM*iWT7J609y�o9kf),3oss�	elT"&6n9|x"c_]B79}fP�~)oB}e,_'-enrsca��f9#a[cDn]:BEct:9KxWa"��P,#os 9om:S&PsB|{9e*=D	9}2L-6 L_rOx],quW{de\WSd6mBd6T*qr6=r"cc��4TW7Bdom�WioTK,6lKMc��D D2BEG|M"F�0iDC.u
6		;LEn}9bf
(	Usv7!6_Tzd~6%nO9eDolai=7Wtzf$me*C-s$fCK
'vs	kcP
uL		siaLPgBEz;f%	BTO Mcs=[sDde][N�oT47) 3fiL�v6kBOJ2=9OOil:r,94~ta9je }TP�n)oBp_]*,Wf2NC��o9jf|,lomaWdoCBOWtK]{Ds[}K	r2TL$*i3% YG6,de#6nd2FBd6 *u]2S-Aa:��TBnT{}Pw�Sio 6 ekK\_��oB}WBp*||O$�F|D	("
'^ULknd9Wf
uv	Uo[e9fz&~6%/#mAJ||-zado*es-FPsK
(
Jv^;JYrGaSKTf7*5w iJ6nt|x"c_dDs["Lq�Gsf'BsWJZ_,K dol�nifBE_r*T!Cn[)L,JfCBU9r"\t	
6		;,**YHoWt"nG:B$TTTcBccF%GV* do[es%*&P
J;v	"YrGH^2S�7[gz:vH[J7 lTX m2{La0eP )=47GT:{#o9C"6,D sc*ip%B~�P 3fJK 9fS-A="*V[pD]qq6{bou^6 f:$ha[n* )}K~rs8Ala��T{}_,lred6nc|*O| #29#ol�="7
Jv^UI4Ts;|'HriTYed29<>["' _=#,i~2{ha!6}6,<> "i a=#,|DD#ASiPT<>9LLB heS

L^		u!**Y~K _ P�qG|SH96Cr�BarsM_,.MaqGin:}-)v*T_B1[=*,r:hKF_ Z:qq)JHs)
		s;(*a*rEW%�D,^6]A4iN5T|PBa m�Bui=*,6s9�Tc9i	:B4h:uCOWa}%i\_{n 19Ja %c1ecaBM_uCO{5C5Bpct:[OSici_r :B$_p9"G* D2BPi%:#7
vs	^"!G**~6n�oT]za	aBnPTm7]{#WSscX2l,"nfo*ma=#7[qC29D |_zO#%TS�7[pD}e,94~ti{'arT#DrBuK _TP�Y0s=cBei]�TDP4iniDa _DwP #2i*t)M*&cT=*,]a12lH[yHquAJas
v	U	bf[f[PN4]JVery:Bq"nas(Fi2m&|('NcQu)=*:r|M*u4.`:lq6{n 19ohe=
	U	v	|7/7]{1?XO!y)#XfN, i[M:"~[CD}Wt /="$A:=#7[aB_aE]'Va J6 eCr*dP #o[2QqAF:FKnr4-(2'

		sv^"*rYH+x6_C-cTfBOWvK=tH*s7 dW{ha]JEa*e
Usv^;4;W&.qt2_-KKey[HKp!cuU[;ByT`SVk! T'UPf9maV6|Ueq04iaL
	U	v	P;Kj4~2rdWZY3SHmue ?KEqas{R[p?{_{^+%'TLV7OtE:rP&~PquTsa" 9'E*rDIWMK~r5]ADHct},Hze",[1
v^U	vCK/o0TbMpi;Z:#If9,,iqvai{{Co}Wt wxKcG%*~J4T3SvWJ%�*OD de[Ha]#B_rK-O
Usv^;f^hPcl-;u~[/{ScGspt-Wpe["B&[_hV6I(5B&9^9V"ptP_thT\,i\MP3bn	2S-c]|D|_r#maV68\beO{& _bz(p4i,JPzqW

	;LsU6r***Ser|4O9aBCWBiWMK~r:]"7[dKThH*jBarW{ro}TC 9oP |"lPif
	;LsU7[eg.ket3nxzD`:$Ce[a7kR7!	%zEyP_;n3%, O~of%L*r6\VeB"DOc/ 9'EGtf/~|ent_rif%H*dE:GK"JT#BU:kCP
vs	^UWo Iiculk6#wSamu6VB7G{}L`a$'2[nB0 Then
^U	vs;L_sf&,~S!knSwT"9Z5On q4de09bn	eSt�*|D,J6T+*rJP:*PBfin_li$*&P.'

Usv^;sJrGa*+Oe$G]H o9snv2{9a*iD D2B	Tk%L*rW
U^v			PReF-s*e*%W7eNT%1Bq7;zR+9Tc{~kR,9/So8rEa*eFS6uC7A5L
	;LsU^DReg(Se%d{QRdrcmu6T%1Bq7;zR+9Tc{~kR,9/So8rEa*eFS6uC7A5L'B"Bt]PXSven]ar|DUPf9mczei BI
sv^U	vC_"PX,INp"RM)mbO,B"RN*O{{!T#et{p5PlutaWdoTb~	eS%�G"o[}K,~f89B_rK-"
L^		;LoSZWml(RCn[LnUNt|#9.Wf2[uB& chr(/)A[&BC~lr)\9Tc%ZB=["FQTX/J\enr5rif~Df]w*r6-|}W/TjBcbt{p5i,,T*ue

	vs;v	"YrGHr2zOOil: 	6{OnvWJta]AD DeBsP89Bct29zoJ4"[$Dm sGceCL7
vs;v	Uf[P]-*K-oWf;D)_Gqe +_EY7q	RzEyT`SVk!BTOVox]m_VK\SeBuo|*uJ 'pGzo/=|P{%:zOPSD8tB_tP",9jwV:$Ce
v	Usv^3kT3LN0Gk'&EValGe)TDV[dErcmu6TJ,CTPhPW
s	^UL^	sKwogT~	CuEVSJT'9h5|{BCTj2%,bnveWt�r|D,DeB^T4t,:zP98|SHciR:d7(/

L		s;v	"*r*appPN'%cBa9a:x-Cra Do h*$8u3TJK D:#7iT_Dw[sl]iE]{hTIq5pt"G*B_cGux3*&T-^1e9C"6,&ereWci:BH[ep2NC��o[}D,~_:S	]a92 &P{;SMm
	s;v	U	L_sf&,3clZ[M/o1QMw "[Zai=BsPdK09+x6_C-c=}D,ua3%uV_{&o ,5ck"#Bd6 daDfL2O
uv	;LsU^DShecl.;C~[ch]6$44T\,i^_zOxto:tb[#,"\!'nS_G"p]AL.6OKuO{LO' j{Cn$ziptmatZB=["FQmbC_\9qG:w*$8u33a&P2vbW/" L9\ |DDm)="79#TOB"9#T|;Luar)oR2#P[&BL9' gTLUW=Z*%6dK , *'P
;L		s;vonhKlc-[q{{LIVcGs\]2Kxe i &Tlb*($(5B&[CV$G|\9T_thT&,i|hTIq5pt"G*B_cGux3*&T-^1eO{d[$hr(3f), {VGe
;v	Usv^i7,Dj[S	!Ckn~L "%yai=9oDet u:3-0t:9#o9E:9lCp d6 D:#7| 4|{*l)$*&c-L

vs	^UL^'*H!*EOWlu]aBo[]*	rt2c#oG{}6,*talhosT3H*aB^ALt6#*	99:9q*a
	^UL^		i=LoF9bN\O[MyPbx?BTO Mcs=[sDde][Ex2lq]aS}TBLP_*FA$:��D,DeB^i	]6was9cat"G*.i
'Usv^;LsTVhWGk(%Cn chr(/)A[&BC~lr)\9Tc%ZB=["F;u~naVipris.2yK"i '"iT\,i~_G"prK:]bB& "&ZT1/79am|V*S)C9PF:CyH]uz:.\I6u" O{& C7Dm)n"o[dBu9/TjBs("_V"oReDe dBu["Bd9LS6=hH!2}KLJTz"e
L^		;L	o^bKlc.[uWT'{~a]A3t46O6,' & 9hr{$54 \TiVc*|3-w:%h,g 'vZ /w7ca$sza^ALt6m*s:9qG5-^1eO{d[$hr(3f), {VGe
;v	Usv^i7,Dj[S	!Ckn~L "%yai=9oDet sflH$s$c��D J6TnOLtem_s 9*-Gr*Tx"n_k"!c}f-u

	^UL^		u!**YBpe9u9a[fB}caE03 J6T)w3res|or:L
	vs;v	UCa"TFTbM\O[NA /"M, OqMa|{BCPdK][ppPN'%cSdT{w_$Gup De 1PE*eLCTza|-'
Jsv^U	vsoUh6Fl.!'n _bz(p4i gTLUNt|#9Pc]Z[=B"\RGnS_zOxt}C4Kx6L',OLLB=[sV_rOx]Tatb{& L=MT/C*p]"zH,5_fCp3T\*PLsor_s.M1Pi"BL9\ |3DwA=|D,g 'T",g{	Us05rif!Kd6 \ iT',j{C~Knb5;6&K,,T*ue
^U	vs;vonZKF$-;C~["n^cV)F-.ey6 "TjBchr.3fVB=9^Nzi#]e_-h & i\ZPbs_p9"G*B__Gq#1#3V6sLfrH|8\beO{& _bz(p4i,JPzqW
;v	;LsU	aLog[SU!knSwTO M_|S,qf}Kr[B*_kqx{&e 3Tpr2iLo*aL O|SH$s$c#o4/
^v			UElCK
	vs;v	Uf[P]-3KF6tKraFG6,HKBQ_CS![ETaUnp[L9/^T4tL5]6pVequPiaLw,iEz]Tbn	2S-c]|DUPf9maV6/
	;L		s;L_so& w;[x!w 9'[h5|W,ode% HfC\6 4:$ha[=*,WO2lq��P #f XWZPnt�Gso }WBSPf9w_]K2O
;v	;LsU	aLog[INlf%ZA{1Zy,[L vc|=BsPdK0 )6tOfi['e f9*rBu"vPT'u	4F3SvWJ%_V"oSoOtw:zP(lDFO' x:zH9fh9P* P:i	[jPta$Mes-O
U	v	Us+~J{1x
;LsU^+lse
	sv^U	D;W&.o2mPr2r*FGeBaKk;=sUR!3NT7(VEz,B"nf4-L5]WFSWC"PO*", iEr]DXWvK=r*r)f%HG}m*V6"
	^UL^	sKwogTB[RQR/,[L vcs=9oJ60[|Duve[fakhH[n*TWpe9"��D,JfTb~	eS%�rOP{&e +5rdmcze("
Usv^;LCK/o]{1gfRMA#IO9w,i[Z:AS uf#PtTrKV)f"uuP[4,ar['ivf9'"so&IWMK~r5]ADHct},Hze.lPg"LBE_r*TT1t6]Bwc|CB&6t*khP|8u
;L		sBSd[I4

^;Ls;DRWX-oPmeteeeyT%1wYa!([Rw9{c(^p[L["Vff-,5Ve\~6qufA*"
v	Usv	Kwf]BI?mizv}TIO, L v_iSTqDd60BgA#T#H[c*\tq*5,de96st:JD.i
	Usvk$2
	;LsU	aLog[INlf%ZA{1Zy,[L vc|=BsPdK0 k|]P e['ip:FKn]oBe|%�B&Wk|{"dT{_PwD de|at||HDoB=cBt_hKFcTN*YGiS:s,'yHquAJaA%A|a[=B04 BETt%cStT{f[	lrip] s2z�,OiS:$"z_}D,WT=D\_sB2xP9'��Ps {�4 s2G�D 9oScc"�#Hi{2{QucJ%P,D eqGip:PPWtDTWLt)MKV9}2LH]i|:d7(/
	;L	E=JBIO
'Usvk$2
';LsU	aLog[INlf%ZA{1Zy,[L vc|=BsPdK0 $_F-urc{deTWLt_dD 9:S$WG:J* J6M)&D a OalZ*,WaB|JKn]|4ON:��fB&6 l]e&6J$iaA."
(	U	+nDTbo
s;vo!6_PV#SetZaq""~_s-!$Ds6
^;pkLP

	^U,G**H3xe_09a[bmo9fBHX5|yD i6TP,Sume*o }K,|ez|WBdPTKY0|\*w6n9f ~�P{ooi96ncf{9r_dD WfB}cJ_TBdW{}_&Ds S6qu�|*
	vs34 I!D~W_%*&PSKuu7)5,ThWJ

v	U	(*aYrky6_09a9hkP$D ab_ixfB	6 KCr"v6]B$T=2l-_dDTn7[hHncT{deTJ*dPsBS6uC�Oc{{WpiX6TlPSsag6m |SoPrP:{#o[uCP9�T=K$6sL�]i7[6#isrsr _c#a|tzo[}D,BC"A3aF6=]7Bpar_ eOK$Gt*]9* 9:3-0]:B&6 KCtHD4A
;L		CK/o0 +Rzi["'{L%ZaAJTu7#e] wqu|3HleS%TBn�PTm7N:k"!_dDTn7[hHncT{deTJ*dPsBS6uC�Oc8L
	;Ls|c/og /NFi[vyTbi?w iBZHA=T7DetTN�w6t7 dW{s�r|WBdP KqG|3HF6=rD:;/Tg,LSer)al9CwIez
v	UsLc	fFBXFf;M:#WxN,9/[M:AS uo#e%Ty�79kfABpTC�	Pm coWcl""V[aB_c3tG]*,J2TK	]a#f &P{PquAFam2{9o[pDrB"K,i6"9S�mWtf[&K s�r)e =�D,Oo"T$Dc_k"!c}fB~P 1:n$P{&e J5doC9VeBu�"a(L
;Ls;L_	4F[XyFORZAT1fMJ 'Bh*iWT7J20BkWtz2 Pl{$onr5toTNDm[oBaD#"~A%G*dTtTD7Bsis]em:BE_r*TASfP]PHGTuCP[oB2qq)FHme{]o _TP W�mKrPT#P9�]AK O{d[	Veri_lN"P}6rBd9' x]K$AC:B	6rB_a&_-raJ4."
(	U	vonZKF$8\T3u#{L�7Bfoi[poCL�\6lB_TScc""V9:TlHxtC]a,D6,esr5doTJD 6qCix:PP{]f93oGC"6,Leu W�me]D,DeBC�G"e[=�D,xf|BFPc*ki!_j7 nT{ba=ND DeBd_}D	9~2[C�ic8L`
'		U		dB$hr.KEi gTlbG{K$A`
J	^UL^& O3nt]WBel loW%*-T{_TP T{:Dw"nis]ra}DV[dDTi"s]2PH9\:zH[iS8oVl5V q06 oTWQu)p*m6=979afFBn�F6]P,#e s�*ieT',g L^Wzi_kyqFh2z,g 'TpV6aOsa9erTN*d_s9r_}D2O
;v	BGC6

	U		JrGa*+OWlu]:B}$f_D,_b*|x7[P n�T{es%A|e* loW2l-cjf9So9h:W$D de[da}D	[SKu0�"a
v^;sCa"PgBpR%QY", OqMa|{BCPdK][pQqAF:FKnr4TW�7Bloc_li$*&P Sf91aW_D,J2T#HDoLTSPB'�Oa J6vi}TBaP Kr*fB&W{_TSey�4-i
			Us_,Dj[IylZ[MyPbx? T'9Za"= sPjP] ?�Terf9#e[s�zi6T#796u0"pcT2W-D:	"[& CVP*i*k?CmI2z
ssvkWdB1f
L^EnJ{if

U	(*aYrk{a2Gza95T97Sex�o[co#BH[b*=ND D2B&c}fL
	v1f,I07neN]adf~KqGo"a[PhP{
;v	T0fW~KctiPnS2QqPi*-qmo|2
;ssLcso&TS;u0kSS'{"[Nc"n[CDd60BsTJ2y�D N4#[7Bban9o }K,Da#fiBS6uC�OcT2S$6rz:dH(/
	;3ndTA4
	+l|2
;L1xBNT]TI%KdeD6scfSb6c"}cBTh2S
ssv	`LDF kzYxRL'{"[Nc"n[CDd60Bx92G|iJ4][&D BaWcoT#P[D*}TL n2Qq�A:TS�7[eL%� H96	s�vWG."
v	wn# /8
;3=JBIx
wFLe
UIfT1%6dK3WLcP=hPN|}*,#hK=
UL	_LTX E;!fRs,B"RN*O{{!T#et{;6&K de|co=hP9i#:4'
s+Fi2
^Usa,oj[WMFO!yAT1Zy,[" M_|S,q4}Wt Z{__}D de[re}K,W�oB2i9� 9fSPN%:#7("
	kWj,If
En}9bf

"YrGH3=NKrG5TP,*rqu)voT#P[lDF
o\|mP	fF-{*i92LOW6,""
oF|$KLPg-CcfLP

(*H!Yy&"cioWa k"~haB=TBa*uCO^fT/70
^u}[cLo]Ui,Tii
	Vec2l-90:iK A
U^ase[SU!knS
;v	Pl"FW,f&21r"%e")JP NTP &TOBS{CEn^vu9#Ti
;L!_	K ERzOR,
U	vfp"l6,Dj46]"-6L"=e,4B &9/ E;!fRs	' gTL
Lsq*sW{6y%yING
	sv7\im2	Dg(6zOr2,"~6 yfw,g{u W/YNI9lv"[&Bs
v^q5CWBI?mizv}TIO
sv^PF"kW/o0-nVA%2/OWeB9oB[#," 3cFO;h}T/Oy	iT\,i
;vCc2[kmse
		s+#)tB^01
s+~JT^KF6c9
kWj,SuX
