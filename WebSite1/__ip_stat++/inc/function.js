//验证码
function reloadcode(){
	var verify=document.getElementById('safecode');
	verify.setAttribute('src','inc/captcha.asp?'+Math.random());
	//这里必须加入随机数不然地址相同我发重新加载
}