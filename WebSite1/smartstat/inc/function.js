//��֤��
function reloadcode(){
	var verify=document.getElementById('safecode');
	verify.setAttribute('src','inc/captcha.asp?'+Math.random());
	//�����������������Ȼ��ַ��ͬ�ҷ����¼���
}