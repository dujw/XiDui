<%
dim rs_message
Set rs_message=conn.execute("select sitename,sitedescription,sitekeywords,kaiguan,guanbi,usertype1,usertype2,usertype3,usertype4,usertype5,usertype6,siteurl,adm_mail,adm_comp,adm_address,adm_name,adm_tel,adm_mob,adm_icp from dk501_bconfig") 
	'网站基本信息
	sitename        = Trim(rs_message("sitename"))
	sitedescription = Trim(rs_message("sitedescription"))
	sitekeywords    = Trim(rs_message("sitekeywords"))	
	kaiguan         = Trim(rs_message("kaiguan"))
	guanbi          = Trim(rs_message("guanbi"))
	type1           = Trim(rs_message("usertype1"))
	type2           = Trim(rs_message("usertype2"))
	type3           = Trim(rs_message("usertype3"))
	type4           = Trim(rs_message("usertype4"))
	type5           = Trim(rs_message("usertype5"))
	type6           = Trim(rs_message("usertype6"))
	
	siteurl         = Trim(rs_message("siteurl"))
	adm_mail        = Trim(rs_message("adm_mail"))
	adm_comp        = Trim(rs_message("adm_comp"))
	adm_address     = Trim(rs_message("adm_address"))
	adm_name        = Trim(rs_message("adm_name"))
	adm_tel         = Trim(rs_message("adm_tel"))
	adm_mob         = Trim(rs_message("adm_mob"))
	adm_icp         = Trim(rs_message("adm_icp"))
set rssetup=nothing
%>