
function do_ping()
{
	if(typeof(pgvMain) != 'function')
	    return;

    // 1. ��ȡdomain/uri/refer
	var domain="";
	var uri="";
	var search="";
	var ref="";
	try
	{
		domain=top.location.hostname;
		uri=top.location.pathname;
		search=top.location.search;
		ref=top.document.referrer;
	}catch(e)
	{
		domain=window.location.hostname;
		uri=window.location.pathname;
		search=window.location.search;
		ref=window.document.referrer;
	}
	domain=domain.toLowerCase();
	uri=uri.toLowerCase();
	search=search.toLowerCase();
	ref=ref.toLowerCase();
	
	/*�ؼ�·��
	 ��Ʒ�������̹ؼ�·������2008.3.1               1)ѡ����Ŀ 2)������Ʒ��Ϣ 3)�����ɹ�
	     my.paipai.com
	 ���ĺ��Ĺ������̹ؼ�·��1.1.1   				1)���� 2)��Ʒ����������� 3)�µ�ҳ�� 4)�µ��ɹ�ҳ�� 5)��CFT���� 6)CFT����ɹ�
	     search.paipai.com search1.paipai.com list.paipai.com auction1.paipai.com shop.paipai.com auction.paipai.com pay.paipai.com
	 �����������µ�(�����¼)�ؼ�·��1.1.2    		1)���� 2)��Ʒ���� 3)�µ�ҳ��
	     search.paipai.com search1.paipai.com list.paipai.com auction1.paipai.com auction.paipai.com
	 �����������µ�(���¼)�ؼ�·��1.1.3     		1)���� 2)��Ʒ���� 3)��¼ 4)�µ�ҳ��
	     search.paipai.com search1.paipai.com list.paipai.com auction1.paipai.com member.paipai.com auction.paipai.com
	 �����µ���CFT����ɹ�(ʵ����Ʒ)�ؼ�·��1.1.4	1)�µ�ҳ�� 2)�µ��ɹ�ҳ�� 3)��CFT���� 4)CFT����ɹ�
	     auction.paipai.com pay.paipai.com
	 �����µ���CFT����ɹ�(������Ʒ)�ؼ�·��1.1.5	1)�µ�ҳ�� 2)CFT����ɹ�
	     auction.paipai.com pay.paipai.com
	*/
	
	/*��ɢ·��
	 miniportal.paipai.com
	 www.paipai.com
	*/
	
	var reg = /^[1-9][0-9]{4,9}\.paipai\.com$/gi;  
	if(uri == "/cgi-bin/sell_type_choose")
	{
		pgvMain("pathtrace", {keyPathTag: "2008.3.1", nodeIndex: 1}); return;
	}
	else if(uri == "/cgi-bin/commodity_type_choose")
	{
		pgvMain("pathtrace", {keyPathTag: "2008.3.1", nodeIndex: 2});return;
	}
	else if(uri == "/cgi-bin/commodity_update_success")
	{
		pgvMain("pathtrace", {keyPathTag: "2008.3.1", nodeIndex: 3});return;
	}
	//���� search.paipai.com  search1.paipai.com list.paipai.com
	else if(domain=="search.paipai.com"||domain=="search1.paipai.com"||domain=="list.paipai.com")
	{
		pgvMain("pathtrace", {keyPathTag: "1.1.1|1.1.2|1.1.3", nodeIndex: "1|1|1", virtualURL: "/", virtualTitle: "����"});  return;
	}
	//��Ʒ���� auction1.paipai.com
	else if(domain=="auction1.paipai.com")
	{
		pgvMain("pathtrace", {keyPathTag: "1.1.1|1.1.2|1.1.3", nodeIndex: "2|2|2", virtualURL: "/", virtualTitle: "��Ʒ����"});  return;
	}
	//�������� shop.paipai.com
	else if(domain=="shop.paipai.com" || (reg != null && reg.test(domain)))
	{
		pgvMain("pathtrace", {virtualDomain: "shop.paipai.com", virtualURL: "/", keyPathTag: "1.1.1", nodeIndex: 2, virtualTitle: "��������"});return;
	}
	//�µ� auction.paipai.com
	else if(uri == "/cgi-bin/commodity_fixup_confirm")
	{
		pgvMain("pathtrace", {keyPathTag: "1.1.1|1.1.2|1.1.3|1.1.4|1.1.5", nodeIndex: "3|3|4|1|1", virtualTitle: "�µ�"});return;
	}
	//�µ��ɹ� pay.paipai.com
	else if(uri == "/cgi-bin/pay_detail_confirm")
	{
		pgvMain("pathtrace", {keyPathTag: "1.1.1|1.1.4", nodeIndex: "4|2", virtualTitle: "�µ��ɹ�"});return;
	}	
	//CFT����ɹ� pay.paipai.com
	else if(uri == "/cgi-bin/pay_success_return" || uri == "/cgi-bin/mass_pay_return")
	{
		pgvMain("pathtrace", {virtualURL: "/cgi-bin/pay_return", keyPathTag: "1.1.1|1.1.4|1.1.5", nodeIndex: "6|4|2", virtualTitle: "����ɹ�"});return;
	}
	//��¼ member.paipai.com
	else if(uri == "/cgi-bin/login_entry")
	{
		pgvMain("pathtrace", {keyPathTag: "1.1.3", nodeIndex: 3, virtualTitle: "��¼"});return;
	}
	//miniportal miniportal.paipai.com
	else if(domain=="miniportal.paipai.com")
	{
		pgvMain("pathtrace", {pathStart: true, useRefUrl: true, override: true, careSameDomainRef: true});return;
	}	
	//��ҳ www.paipai.com
	else if(domain=="www.paipai.com" && uri=="/")
	{
		pgvMain("pathtrace", {pathStart: true, useRefUrl: true, override: true, careSameDomainRef: true, virtualTitle: "������ҳ"});return;
	}
			
    pgvMain();
}

//s=(new Date()).getTime();
do_ping();
//e=(new Date()).getTime();
//alert('pgv:'+(e-s));
