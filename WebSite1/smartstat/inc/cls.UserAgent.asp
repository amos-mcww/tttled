<%
class CheckUserAgent
	public vOs
	public vSoft

	public function execute(strUA)
	  vOs=trim(GetOs(GetWin(strUA)))
	  vSoft=trim(GetSoft(strUA))

	  if vOs="" and instr(vSoft,"Konqueror") then vOs="Linux"
	  if instr(vSoft,"Mozilla") and instr(strUA,"compatible") then vSoft=""
		select case vOs
		case "Windows NT 5.0"
			vOs = "Windows 2000"
		case "Windows NT 5.1"
			vOs = "Windows XP"
		case "Windows NT 5.2"
			vOs = "Windows Server 2003"
		case "Windows NT 6.0"
			vOs = "Windows Vista|2008"
		case "Windows NT 6.1"
			vOs = "Windows 7|2008 R2"
		case else
		vOs = vOs
		end select
	end function
	
	private function GetOs(strUA)
	  dim regEx ,match,matches,maxlong
	  GetOs=""
	  maxlong=0
	  Set regEx = New RegExp
	  regEx.Pattern = "(Windows|Mac |Mac_|Win|PPC|Linux|unix|SunOS|BSD)[^;\(\)\[]{0,20}"
	  regEx.IgnoreCase = True
	  regEx.Global = True
	  Set Matches = regEx.Execute(strUA)
	  For Each Match in Matches
	    if match.length>maxlong then
			maxlong=match.length
			GetOs=match.value
		end if
	  Next
	end function
	
	private function GetSoft(strUA)
	  dim regEx ,match,matches,maxlong
	  GetSoft=""
	  Set regEx = New RegExp
	  regEx.Pattern = "(Konqueror|Opera|Safar|Firebird|NetCaptor|MSN |Netscape|MSIE|MyIE|OmniWeb|AOL|WebTV|iCab|Mozilla)[\d\/]?\d*\.?\d*\.*\d*[^;\(\)\[]*"
	  regEx.IgnoreCase = True
	  regEx.Global = True
	  Set Matches = regEx.Execute(strUA)
	  For Each Match in Matches
		if instr(GetSoft,"Mozilla") then
			GetSoft = match.value
		else
			GetSoft = GetSoft & "," & match.value
		end if
	  Next
	  GetSoft=replace(GetSoft,"/"," ")
	end function

	private function GetWin(strUA)
	  dim regEx
	  Set regEx = New RegExp
	  regEx.Pattern = "Win\s?(\d+|NT)"
	  regEx.IgnoreCase = True
	  regEx.Global = True
	  GetWin=regEx.Replace (strUA,"Windows $1")
	end function

end class
%>