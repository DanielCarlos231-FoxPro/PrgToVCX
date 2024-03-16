*PRVCPRGPAT = "C:\ArtSystem\Testes\PrgClass.PRG"
*PRVC_CLASS = "PrgClass1"

PRVCPRGPAT = "C:\ArtSystem\Thor\Thor\Tools\Components\ExcelXML\ExcelXML-master\ExcelXML.prg"
PRVC_CLASS = "ExcelXML"
Clear 

*PrgToVcx = CreateObject("PrgToVcx")
*ClassArr = PrgToVcx.APrgClass(PRVCPRGPAT, PRVC_CLASS)

*Set Step On 
*Return 

PrgToVcx = CreateObject("PrgToVcx")
*PrgToVcx.ClassFromCode(PRVCPRGPAT, PRVC_CLASS, PRVC_CLASS+"_VClass", "C:\ArtSystem\Testes\"+PRVC_CLASS+"_VClassLib")
PrgToVcx.ClassFromCode(PRVCPRGPAT, PRVC_CLASS, PRVC_CLASS, "C:\ArtSystem\Testes\"+PRVC_CLASS)

DEFINE CLASS PrgToVcx AS CUSTOM
	
	* ClassMembers[1,1] = "Class"
	* ClassMembers[1,2] = "Name"
	* ClassMembers[1,3] = "Value"
	* ClassMembers[1,4] = "Type"
	* ClassMembers[1,5] = "Visibility"
	Hidden Array ClassMembers(1,5)
*	Dimension ClassMembers[1,5]
	
	
	PROCEDURE INIT

	ENDPROC

	PROCEDURE DESTROY

	ENDPROC
	
	Hidden Procedure CleanClassMembers
		
		Dimension This.ClassMembers(1,5)
		Asin(This.ClassMembers,1,1)
		Asin(This.ClassMembers,1,2)
		Asin(This.ClassMembers,1,3)
		Asin(This.ClassMembers,1,4)
		Asin(This.ClassMembers,1,5)
	
	Endproc
	
	Procedure APrgClass
		Lparameters plc_ClassPath, plc_ClassName

		&& Limpar Class Member Array
		This.CleanClassMembers()

		* la_ClassMemb[1,1] = "Class"
		* la_ClassMemb[1,2] = "Name"
		* la_ClassMemb[1,3] = "Value"
		* la_ClassMemb[1,4] = "Type"
		* la_ClassMemb[1,5] = "Visibility"
*		Dimension pla_ClassMemb(1,5)
		
		Local lc_Class, lo_Class, lc_ClassPath
		lc_Class 		= plc_ClassName
		lc_ClassPath	= plc_ClassPath
		
		lo_Class 		= NewObject(lc_Class		, lc_ClassPath		) && Object Class 
		
		&& Properties
		This.AddProperties(lo_Class,"G")	&& Public
		This.AddProperties(lo_Class,"P")	&& Protected
		This.AddProperties(lo_Class,"H")	&& Hidden
		
		&& Methods
		This.AddMethods(lo_Class,"G")		&& Public
		This.AddMethods(lo_Class,"P")		&& Protected
		This.AddMethods(lo_Class,"H")		&& Hidden
		
		Dimension pla_ClassMemb[Alen(This.ClassMembers,1),5]
	*	Acopy(This.ClassMembers, pla_ClassMemb)
	*	Set STEP ON 
			
	*	This.CleanClassMembers()
			
	EndProc
	
	Hidden Procedure AddProperties
		Lparameters plo_Class, plc_Flag
		
		Local lc_Flag
		plc_Flag = evl(plc_Flag,"")
		
		DO CASE
			CASE plc_Flag = "G"
				lc_Flag = "Public"
			
			CASE plc_Flag = "P"
				lc_Flag = "Protected"
			
			CASE plc_Flag = "H"
				lc_Flag = "Hidden"
					
			OTHERWISE
				plc_Flag = "G"
				lc_Flag = "Public"
				
		EndCase
		
		Local array la_Proper(1)
		
		AMembers(la_Proper, plo_Class,0, plc_Flag)
		
		Local x, ln_TotMerbers
		x = 1
		ln_TotMerbers = Iif(Alen(This.ClassMembers,1) > 1, Alen(This.ClassMembers,1), 0)
		
		For x = 1 to Alen(la_Proper,1)
			
			If not PemStatus(plo_Class.Name,la_Proper[x],5)
				loop
			Endif
			
			ln_TotMerbers = ln_TotMerbers + 1
			Dimension This.ClassMembers(ln_TotMerbers,5)
			This.ClassMembers[ln_TotMerbers,1] = plo_Class.Name
			This.ClassMembers[ln_TotMerbers,2] = la_Proper[x]
			This.ClassMembers[ln_TotMerbers,3] = GetPem(plo_Class.Name, la_Proper[x])
			This.ClassMembers[ln_TotMerbers,4] = "Property"
			This.ClassMembers[ln_TotMerbers,5] = lc_Flag
			
		EndFor
		
	Endproc
	
	Hidden Procedure AddMethods
		Lparameters plo_Class, plc_Flag
		
		Local lc_Flag
		plc_Flag = evl(plc_Flag,"")
		
		DO CASE
			CASE plc_Flag = "G"
				lc_Flag = "Public"
				
			CASE plc_Flag = "P"
				lc_Flag = "Protected"
			
			CASE plc_Flag = "H"
				lc_Flag = "Hidden"
					
			OTHERWISE
				plc_Flag = "G"
				lc_Flag = "Public"
				
		EndCase
		
		
		Local array la_Method(1)
		
		AMembers(la_Method, plo_Class,1, plc_Flag)
		
		Local x, ln_TotMerbers
		x = 1
		ln_TotMerbers = Iif(Alen(This.ClassMembers,1) > 1, Alen(This.ClassMembers,1), 0)
		
		For x = 1 to Alen(la_Method,1)
			
			If not PemStatus(plo_Class.Name,la_Method[x,1],5)
				loop
			Endif
		
			If InList(PemStatus(plo_Class.Name,la_Method[x,1],3),"Property")
				loop
			Endif
			
			ln_TotMerbers = ln_TotMerbers + 1
			Dimension This.ClassMembers(ln_TotMerbers,5)
			This.ClassMembers[ln_TotMerbers,1] = plo_Class.Name
			This.ClassMembers[ln_TotMerbers,2] = la_Method[x,1]
			This.ClassMembers[ln_TotMerbers,3] = GetPem(plo_Class.Name, la_Method[x,1])
			This.ClassMembers[ln_TotMerbers,4] = PemStatus(plo_Class.Name,la_Method[x,1],3)
			This.ClassMembers[ln_TotMerbers,5] = lc_Flag
		
		EndFor
		
	Endproc
	
	Procedure ClassFromCode
		Lparameters progName, className, newClassName, newLibraryName
	
		IF NOT File(progName) THEN
			Return .F.
		EndIf
	    
	    Local  oPrgClass, oVcxClass
		oPrgClass = NewObject(className,progName)	
		
		* Here is the trick
		* Open any empty custom class from a vcx
		Modify Class "myCustom" OF "C:\Test.vcx" AS "Custom" NOWAIT 
		objectCounter = ASelObj(arrayObjects,1)
		oVcxClass = arrayObjects[1]
	*	oVcxClass = CreateObject("Custom")
	*	Set Step On 
		
		This.APrgClass(progName, className)
		Acopy(This.ClassMembers, arrayMembers)
		membersCount = Alen(arrayMembers,1)
		
		FOR counter = 1 to membersCount
			memberName = arrayMembers(counter,2)
			memberValu = arrayMembers(counter,3)
			memberType = arrayMembers(counter,4)
			memberVisi = arrayMembers(counter,5)
			
			DO CASE
				Case Lower(memberVisi) = "public"
					memberVisi = 1
				Case Lower(memberVisi) = "protected"
					memberVisi = 2
				Case Lower(memberVisi) = "hidden"
					memberVisi = 3	
					
				OTHERWISE

			ENDCASE

			
			DO CASE
				CASE memberType = "Property"
					prgProperty = "oPrgClass." + memberName
					vcxProperty = "oVcxClass." + memberName
				
				*	IF NOT InList(memberName,"CONTROLS","OBJECTS","LEFT","TOP") THEN  && Banned properties
					If not PemStatus(oVcxClass,memberName,1) && Avoid Read-Only properties
						
						&& Public Properties
						&& Todo - check if it is possible to get the values of the hidden properties
						IF Type(vcxProperty)=="U" Then
							oVcxClass.AddProperty(memberName, memberValu, memberVisi)
						EndIf
						
					ENDIF
				
				CASE InList(memberType,"Method", "Event") 
					methodCode = memberValu
									
					IF NOT Empty(methodCode) THEN
					*	oVcxClass.WriteMethod(memberName,methodCode, .t.,memberVisi)
					Else	
						methodCode = this.ExtractMethodCode(className, memberName, progName)
					*	oVcxClass.WriteMethod(memberName, methodCode, .t., memberVisi)
					EndIf
					
					methodCode = This.TreatCode(methodCode) && Em testes
					oVcxClass.WriteMethod(memberName,methodCode, .t.,memberVisi)
				Otherwise
				
			EndCase
			
		*	If memberName = "INIT"
		*		Set Step On 
		*	Endif
			
		ENDFOR
		* Save the visual class
		oVcxClass.SaveAsClass(newLibraryName,newClassName)
		
		* Close the empty custom Opened class
	*	Keyboard '{ESC}'
	*	Keyboard '{LEFTARROW}'
	*	Keyboard '{ENTER}'  
	Endproc

	Procedure ExtractClassCode
		lPARAMETERS lcClassName, lcPRGFilePath
		
		LOCAL lcPRGCode, lnStartPos, lnEndPos, lcClassCode, lnOccur
		lcClassCode = ""

		* Check if the PRG file exists
		IF NOT FILE(lcPRGFilePath)
			Return lcClassCode
		EndIf
		
	    * Read the contents of the PRG file into a string
	    lcPRGCode = FILETOSTR(lcPRGFilePath)
		
	    * Find the start and end positions of the class definition
	    lnStartPos 	= ATC("DEFINE CLASS " + Upper(lcClassName), Upper(lcPRGCode))
	    lnOccur 	= Occurs("DEFINE CLASS ", Upper(Substr(lcPRGCode,1,lnStartPos)))+1
	    
	    * Check if the class is found
	    IF lnStartPos < 1
	    	Return lcClassCode    
	    ENDIF
	
		* Locate the end of the class definition
        lnEndPos = ATC("ENDDEFINE", Upper(lcPRGCode),lnOccur)+9
        
        * Check if the end of the class definition is found
        IF lnEndPos > 0
            lcClassCode = SUBSTR(lcPRGCode, lnStartPos, lnEndPos - lnStartPos)
        ENDIF
		
	*	? lnStartPos, lnEndPos
		RETURN lcClassCode
	Endproc

	Procedure ExtractMethodCode
		LPARAMETERS lcClassName, lcMethodName, lcPRGFilePath

		* Check if the PRG file exists
		LOCAL lcMethodCode, lcPRGCode, lnOccur
		lcMethodCode = ""
		IF !FILE(lcPRGFilePath)
			Return lcMethodCode
		EndIf
		
		* Extract Class From file
		lcPRGCode = this.ExtractClassCode(lcClassName, lcPRGFilePath)
			
		* Locate the method definition in the PRG file
		nOccur		 	= 0
		lnMethodStart 	= This.GetStartPos(lcPRGCode, lcMethodName)
	*	lnOccur 		= This.GetOccurrences(lcPRGCode,lnMethodStart)
		
		If lnMethodStart < 1
			Return lcMethodCode
		EndIf
		
		lnMethodEnd = This.GetEndPos(lcPRGCode, lnMethodStart)
		
		IF lnMethodEnd > 0
			lcMethodCode = SUBSTR(lcPRGCode, lnMethodStart, lnMethodEnd - lnMethodStart)
        ENDIF
		
		lcMethodCode = Strtran(lcMethodCode,"PROCEDURE","",1,1,1)
		lcMethodCode = Strtran(lcMethodCode,"FUNCTION","",1,1,1)
		lcMethodCode = Strtran(lcMethodCode,lcMethodName,"",1,1,1)
		lcMethodCode = Alltrim(lcMethodCode)
		
		If Left(lcMethodCode,1)="("
			lcMethodCode = Strtran(lcMethodCode, StrExtract(lcMethodCode,"(",")",1,4), StrExtract(lcMethodCode,"(",")",1),1,1,1)
			lcMethodCode = "Lparameters " + lcMethodCode
		endif
	
		Return lcMethodCode
	Endproc
	
	 Hidden Procedure TreatCode
		Lparameters plc_StrCode
			
		Local 	LOCCSTRCOD AS String ,;
				LOCCFRSLIN AS String ,;
				LOCNTOTTAB AS Integer,;
				LOCC_ENTER AS String ,;
				LOCC___TAB AS String 
				 
				
		LOCCSTRCOD = Iif(Type('plc_StrCode') = 'C', plc_StrCode, "")	
		LOCNTOTTAB = 0
		LOCC___TAB = Chr(9)
		LOCC_ENTER = Chr(13)+Chr(10)
		
		&& Get First line
		Local x, lc_Line, lc_String, ArrLin(1)
		ALines(ArrLin, LOCCSTRCOD)

		x = 1
		lc_Line = ""
	*	Set Step On 
		For x = 1 to Alen(ArrLin,1)
			
			lc_Line = lc_Line + ArrLin[x]
			
			If Right(Alltrim(ArrLin[x]),1) == ";"
				lc_Line = Rtrim(Rtrim(lc_Line),1,";")
				Loop
			Endif
			
			If GetWordCount(lc_Line) == 0 or (At("lparam", Lower(lc_Line)) > 0 and Substr(lc_Line,1,1) <> LOCC___TAB)
				lc_Line = ""
				Loop 
			EndIf
			
			LOCCFRSLIN = lc_Line
			lc_Line = ""
			Exit 
			
		EndFor
		
		&& Count TABs in the First Line
		Local y, lc_Char
		For y = 1 to Len(LOCCFRSLIN)
			lc_Char = Substr(LOCCFRSLIN, y, 1)
			
			If IsAlpha(lc_Char)
				Exit 
			Endif
			
			If  Asc(lc_Char) = 9 && TAB
				LOCNTOTTAB = LOCNTOTTAB + 1
			Endif
			
		Endfor
		
		&& Remove TABs from the begining of all Lines
		Local z, lc_Line, lc_String
		lc_Line = ""
		lc_String = ""
		For z = 1 to Alen(ArrLin,1)
			
			lc_Line = ArrLin[z]
			
			If Empty(lc_Line) and Empty(lc_String)
				loop
			EndIf
			
			lc_Line = Iif(!Empty(lc_Line), lc_Line, "")
			lc_Line = Strtran(lc_Line, Padl("",LOCNTOTTAB, LOCC___TAB),"",1,1,1)
			lc_String = lc_String +Iif(z > 1 AND z < Alen(ArrLin,1) and !Empty(lc_String), LOCC_ENTER, "")+ lc_Line
			
		EndFor
		
		LOCCSTRCOD = lc_String	 	
		Return LOCCSTRCOD	
	EndProc

	Hidden Procedure GetStartPos
		Lparameters plc_PrgCode, plc_Name
		
		Local lnMethodStart
		lnMethodStart = ATC("PROCEDURE " + Upper(plc_Name), Upper(plc_PrgCode)) 
		
		If lnMethodStart = 0
			lnMethodStart = ATC("FUNCTION " + Upper(plc_Name), Upper(plc_PrgCode)) 
		Endif
		
		Return lnMethodStart
	EndProc
	
	Hidden Procedure GetEndPos
		Lparameters plc_PrgCode, pln_Start
		
		Local lnMethodEnd, lc_MethodType
		
		lc_MethodType = Upper(GetWordNum(Substrc(plc_PrgCode, pln_Start,Len(plc_PrgCode)),1))
		
		If lc_MethodType == "PROCEDURE"
			lnMethodEnd = ATC("ENDPROC", Substr(Upper(plc_PrgCode),pln_Start,Len(plc_PrgCode)), 1)
		Else
			lnMethodEnd = ATC("ENDFUNC", Substr(Upper(plc_PrgCode),pln_Start,Len(plc_PrgCode)), 1)
		EndIf
		
		If lnMethodEnd>0
			lnMethodEnd = lnMethodEnd + pln_Start - 1
		Endif
		
		Return lnMethodEnd
	Endproc
	
	Hidden Procedure GetOccurrences
		Lparameters plc_PrgCode, pln_Start
		
		Local lnOccur
		
		
		If Upper(GetWordNum(Substrc(plc_PrgCode, pln_Start,Len(plc_PrgCode)),1)) = "PROCEDURE"
			lnOccur = Occurs("PROCEDURE ", Upper(Substr(plc_PrgCode,1,pln_Start)))
		ELSE
			lnOccur = Occurs("FUNCTION ", Upper(Substr(plc_PrgCode,1,pln_Start)))
		EndIf
		lnOccur = lnOccur + 1
		
		Return lnOccur
	Endproc
	
	PROCEDURE ERROR(NERROR, CMETHOD, NLINE)

	ENDPROC

ENDDEFINE


