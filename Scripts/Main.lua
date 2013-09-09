require "luacom"
require "Config"

g_DM = nil

function main()
	InitialDM()

	local dmRet = g_DM:UseDict(FONT_LIB_NUMBER_INDEX)
	if dmRet == 0 then
		error("DM:UseDict " .. tostring(FONT_LIB_NUMBER_INDEX))
	end

	local pos = g_DM:Ocr(299,717,344,728,"ffffff-000000",1.0)
	print("Position:",pos)

	local dmRet = g_DM:UseDict(FONT_LIB_CHS_INDEX)
	if dmRet == 0 then
		error("DM:UseDict " .. tostring(FONT_LIB_CHS_INDEX))
	end

	local city = g_DM:Ocr(258,715,297,731,"ffffff-000000",1.0)
	print("City:",city)

	assert(city == "����ʡ")

	UninitialDM()
end

function InitialDM()
	-- ������Į���
	g_DM = luacom.CreateObject("dm.dmsoft")

	-- ���ô�Į���·��
	local dmRet = g_DM:SetPath(DM_PATH)
	if dmRet == 0 then
		error("DM:SetPath")
	end

	-- �󶨵�Ӧ�ó���Ĵ��ھ��
	local hwnd = g_DM:GetMousePointWindow()
	dmRet = g_DM:BindWindow(hwnd,"gdi","windows","windows",0)
	if dmRet == 0 then
		error("DM:BindWindow")
	end

	-- ����ʹ�õ����ֿ�
	dmRet = g_DM:SetDict(FONT_LIB_CHS_INDEX,FONT_LIB_CHS_FILE)
	if dmRet == 0 then
		error("DM:SetDict " .. FONT_LIB_CHS_FILE)
	end

	dmRet = g_DM:SetDict(FONT_LIB_NUMBER_INDEX,FONT_LIB_NUMBER_FILE)
	if dmRet == 0 then
		error("DM:SetDict " .. FONT_LIB_NUMBER_FILE)
	end
end

function UninitialDM()
	local dmRet = g_DM:UnBindWindow()
	if dmRet == 0 then
		error("DM:UnBindWindow")
	end
end

main()
