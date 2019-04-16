############################
# Large screen monitor for #
# real business handling   #
# pip install pywin32      #
# pip install 
############################
import math
import win32api
import win32gui
import pywinauto
import win32con
import time
import threading
import pywintypes
from pywinauto.application import Application

# 地市/营业厅-终端IP映射关系
cities = {
	'济南市市中区共青团路直营店': '134.33.107.87',
	'青岛-10月19日后提供': '10.19.247.11',
	'淄博市临淄区晏婴路营业厅': '134.35.170.70',
	'枣庄市市中区文化西路营业厅': '134.45.208.88',
	'东营-东城－运河路营业厅': '134.44.104.117',
	'烟台市-芝罘区-西大街营业厅': '134.37.101.151',
	'潍坊市高新区东方路营业厅': '134.38.101.33',
	'济宁-高新-洸河路营业厅': '134.39.178.140',
	'泰安-市区-东岳大街旗舰店': '134.40.109.125',
	'威海-高区-古寨营业厅': '134.49.101.92',
	'日照-莒县－振兴路营业厅': '134.46.168.159',
	'莱芜市莱城区胜利路营业厅': '134.47.111.161',
	'临沂兰山区金雀山路营业厅': '134.41.101.57',
	'德州-东风中路676号移动营业厅': '134.36.101.32',
	'聊城市开发区东昌东路营业厅': '134.48.113.21',
	'滨州市黄河二路营业厅': '134.43.1.136',
	'菏泽市开发区中华东路营业厅': '134.42.214.7'
}
# 监视窗体摆放位置映射表
slots = {
	'济南市市中区共青团路直营店': 1,
	'青岛-10月19日后提供': 2,
	'淄博市临淄区晏婴路营业厅': 3,
	'枣庄市市中区文化西路营业厅': 4,
	'东营-东城－运河路营业厅': 5,
	'烟台市-芝罘区-西大街营业厅': 6,
	'潍坊市高新区东方路营业厅': 7,
	'济宁-高新-洸河路营业厅': 8,
	'泰安-市区-东岳大街旗舰店': 9,
	'威海-高区-古寨营业厅': 10,
	'日照-莒县－振兴路营业厅': 11,
	'莱芜市莱城区胜利路营业厅': 12,
	'临沂兰山区金雀山路营业厅': 13,
	'德州-东风中路676号移动营业厅': 14,
	'聊城市开发区东昌东路营业厅': 15,
	'滨州市黄河二路营业厅': 16,
	'菏泽市开发区中华东路营业厅': 17
}
# 终端IP-账户映射表
accounts = {
	'134.33.107.87': 'jnradmin',
	'10.19.247.11': 'admin',
	'134.35.170.70': 'admin',
	'134.45.208.88': 'admin',
	'134.44.104.117': 'admin',
	'134.37.101.151': 'ytadmin',
	'134.38.101.33': 'dfl002',
	'134.39.178.140': 'jnywtc',
	'134.40.109.125': 'taradmin',
	'134.49.101.92': 'admin',
	'134.46.168.159': 'jxyd',
	'134.47.111.161': 'lwRadmin',
	'134.41.101.57': 'jinqueshan',
	'134.36.101.32': 'dezhou',
	'134.48.113.21': 'lcdcdlyyt',
	'134.43.1.136': 'admin',
	'134.42.214.7': 'hezeadmin'
}
# 终端IP-密码映射表
passwds = {
	'134.33.107.87': 'Jn@1234567',
	'10.19.247.11': 'Sdyd60_csp',
	'134.35.170.70': 'lzyy03',
	'134.45.208.88': 'zzwhxjk',
	'134.44.104.117': 'admin1',
	'134.37.101.151': 'YTzwzx@2018',
	'134.38.101.33': 'Dfl2018',
	'134.39.178.140': 'jnywtc@0537',
	'134.40.109.125': 'TAmobile2018',
	'134.49.101.92': 'Admin123',
	'134.46.168.159': 'abc123456',
	'134.47.111.161': 'LWslu99',
	'134.41.101.57': 'Aa123456',
	'134.36.101.32': 'Xin123123',
	'134.48.113.21': 'Ab123321',
	'134.43.1.136': 'admin123',
	'134.42.214.7': 'HEzetest_2018'
}
# 终端IP-端口映射表
ports = {
	'134.33.107.87': '4899',
	'10.19.247.11': '4899',
	'134.35.170.70': '4899',
	'134.45.208.88': '4899',
	'134.44.104.117': '4899',
	'134.37.101.151': '4899',
	'134.38.101.33': '4899',
	'134.39.178.140': '4899',
	'134.40.109.125': '4899',
	'134.49.101.92': '4899',
	'134.46.168.159': '4899',
	'134.47.111.161': '8600',
	'134.41.101.57': '4899',
	'134.36.101.32': '4889',
	'134.48.113.21': '4899',
	'134.43.1.136': '4899',
	'134.42.214.7': '4899'
}
# 获取Radmin管理窗口
# 查找Radmin管理窗口句柄，无论是否隐藏
hwnd = win32gui.FindWindow(None, "Radmin Viewer")
if hwnd == 0 :
	# 若句柄不可用，说明Radmin未启动，则直接启动
	app = Application().start('d:\Program Files (x86)\Radmin Viewer 3\Radmin.exe')
	# 等待窗口可用
	app[u'Radmin Viewer'].wait('enabled', timeout = 3)
	# 重新获取句柄
	hwnd = win32gui.FindWindow(None, "Radmin Viewer")
else :
	# 若句柄可用，说明管理窗口存在，直接依附
	app = Application().connect(handle = hwnd)
# 显示窗口，无论是否隐藏
win32gui.ShowWindow(hwnd, 8)
# 全局关闭监视窗口的工具栏
win32gui.SetActiveWindow(hwnd)
try:
	app[u'Radmin Viewer'].menu_select('Tools->Options...')
except pywinauto.base_wrapper.ElementNotEnabled :
	app.top_window()['CloseButton'].click()
	app[u'Radmin Viewer'].menu_select('Tools->Options...')
except:
	print("")
app[u'Options'].wait('visible', timeout = 3)
app[u'Options'].TreeView.select('\\Remote Screen Options')
app[u'Options']['Show toolbarCheckBox'].uncheck()
app[u'Options']['OKButton'].click()

# 线程同步锁
lock = threading.Lock()

# 连接函数
def connect(key, value, show_type):
	# 判断key是否存在如果存在置前
	# 第二个条件用于规避pywinauto按照title匹配窗体时匹配错误的情形
	if show_type == 'split' and app[key].exists() and app[key].window_text() == key:
		hwnd = app[key].handle
		# 窗体前置
		win32gui.SetForegroundWindow(hwnd)
		return
	elif show_type == 'shuffle' and app['* ' + key].exists() and app['* ' + key].window_text() == '* ' + key:
		hwnd = app['* ' + key].handle
		setWndCoords(hwnd, show_type)
		win32gui.SetForegroundWindow(hwnd)
		return
	# 加锁
	lock.acquire()
	# 打开新建连接对话框
	if not app[u'Radmin Viewer'].is_enabled():
		app.top_window().close()
	# 选择菜单
	app[u'Radmin Viewer'].menu_select('Connection->Connect To...')
	# 等待对话框可见
	app[u'Connect to'].wait('visible', timeout = 3)
	# 连接对话框
	wndConnTo = app[u'Connect to']
	# 选择连接模式为仅查看
	wndConnTo.ComboBox.select(u'View Only')
	# 输入被监视终端IP
	wndConnTo[u'IP address or DNS nameEdit'].wait('exists', 3)
	wndConnTo[u'IP address or DNS nameEdit'].wait('visible', 3)
	wndConnTo[u'IP address or DNS nameEdit'].wait('enabled', 3)
	wndConnTo[u'IP address or DNS nameEdit'].wait('ready', 3)
	wndConnTo[u'IP address or DNS nameEdit'].set_edit_text(value)
	wndConnTo[u':Edit'].wait('exists', 3)
	wndConnTo[u':Edit'].wait('visible', 3)
	wndConnTo[u':Edit'].wait('enabled', 3)
	wndConnTo[u':Edit'].wait('ready', 3)
	wndConnTo[u':Edit'].set_edit_text(ports[value])
	# 安全考虑，不添加到电话簿
	wndConnTo[u'Add to the phonebook'].uncheck()
	# 左侧导航树选择远程屏幕
	wndConnTo.TreeView.select('\\Remote Screen')
	# 设置色彩深度
	wndConnTo[u'24 bits'].click()
	# 选择拉伸模式
	wndConnTo[u'Stretch'].click()
	# 确定后建立连接
	wndConnTo[u'OKButton'].click()
	# 等待弹出账号/密码验证窗口,或弹出连接状态提示
	loopCount = 0
	while True:
		# 情形1、正常弹出验证窗口
		if app[u'Radmin security: ' + value].exists() and app[u'Radmin security: ' + value].is_visible() and app[
			u'Radmin security: ' + value].is_enabled():
			# 填写用户名，有时不需要填写用户名
			if app[u'Radmin security: ' + value][u'User name :Edit'].exists() and app[u'Radmin security: ' + value][
				u'User name :Edit'].is_visible() and app[u'Radmin security: ' + value][u'User name :Edit'].is_enabled():
				app[u'Radmin security: ' + value][u'User name :Edit'].set_edit_text(accounts[value])
			# 填写密码，任何情况下都有
			if app[u'Radmin security: ' + value][u'Enter password: Edit'].is_enabled():
				app[u'Radmin security: ' + value][u'Enter password: Edit'].set_edit_text(passwds[value])
			# 提交按钮，提交了才能退出循环
			if app[u'Radmin security: ' + value][u'OK'].is_enabled():
				app[u'Radmin security: ' + value][u'OK'].click()
				break
		# 情形2、弹出连接消息窗口，可能超时、服务端异常引起
		elif app[u'Connection info'].exists():
			# 连接消息窗口可能有瞬间存在的可能，随后就消失，造成is_visible和is_enabled异常
			try:
				# 若不存在，则继续循环
				if not app[u'Connection info'].is_visible() or not app[u'Connection info'].is_enabled():
					continue
			except pywinauto.findbestmatch.MatchError as e:
				continue
			# 情形2.1、连接失败窗口有时候Stop按钮不存在，但exists()返回True，可能是Windows或Radmin的bug
			# 需要增加死循环检测
			if app[u'Connection info'][u'Stop'].exists():
				try:
					# Stop有瞬间存在的可能
					if not app[u'Connection info'][u'Stop'].is_visible() or not app[u'Connection info'][u'Stop'].is_enabled():
						continue
				except pywinauto.findbestmatch.MatchError as e:
					# 检测死循环次数，允许3次内循环
					loopCount = loopCount + 1
					if loopCount < 3:
						continue
				# Stop确实存在，则关闭、退出
				if loopCount < 3:
					app[u'Connection info'][u'Stop'].click()
					lock.release()
					return
			# 情形2.2、连接已失败，关闭消息窗口，并退出循环
			if app[u'Connection info'][u'Close'].exists() and app[u'Connection info'][u'Close'].is_visible() and \
					app[u'Connection info'][u'Close'].is_enabled():
				app[u'Connection info'][u'Close'].click()
				lock.release()
				return
			# 情形2.3、连接已失败，无Close按钮，最小化的一个窗口（只有标题）
			else:
				app[u'Connection info'].close()
				lock.release()
				return
		time.sleep(1)
	# 等待出现认证失败/监视器窗口
	while True:
		time.sleep(1)
		# 情形1、认证失败窗口
		if app[u'Connection info'].exists() and app[u'Connection info'].is_visible() and app[u'Connection info'].is_enabled():
			if app[u'Connection info'][u'Close'].exists() and app[u'Connection info'][u'Close'].is_visible() and app[u'Connection info'][u'Close'].is_enabled():
				app[u'Connection info'][u'Close'].click()
			else:
				app[u'Connection info'].close()
			break
		# 情形2、监视器窗口
		elif app[value + ' - View Only'].exists() and app[value + ' - View Only'].is_visible() and app[value + ' - View Only'].is_enabled():
			# 计算监视窗口宽度．若小于200像素则为连接等待状态，需要等待监视窗口正常打开
			left, top, right, bottom = win32gui.GetWindowRect(app[value + ' - View Only'].handle)
			if right - left > 200:
				break

	# 标题栏菜单有时候加载延迟,建议等待1-2秒
	time.sleep(2)
	# 查找监视窗口句柄
	hwnd = app[value + ' - View Only'].handle
	# 设置窗口标题
	if show_type == 'shuffle':
		win32gui.SetWindowText(hwnd, '* ' + key)
	elif show_type == 'split':
		win32gui.SetWindowText(hwnd, key)
	# 关闭2号屏幕手写屏,可能会和其他线程产生鼠标键盘的争用
	app.window(handle = hwnd).click_input(coords=(13, 13))
	app.window(handle = hwnd).type_keys('R')
	app.window(handle = hwnd).type_keys('{ENTER}')
	# 远程窗口大小、位置设置
	fullX = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
	fullY = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
	if show_type == 'split':
		scrX = fullX / 5
		scrY = fullY / 5
		coordX = (slots[key] - 1) % 5 * scrX
		coordY = math.floor((slots[key] - 1) / 5) * scrY
		win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, int(coordX), int(coordY), int(scrX), int(scrY), win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER | win32con.SWP_SHOWWINDOW)
		win32gui.SetForegroundWindow(hwnd)
	elif show_type == 'shuffle':
		setWndCoords(hwnd, show_type)
		time.sleep(1) # 修复窗体覆盖时自动回位的问题
		setWndCoords(hwnd, show_type) # 修复窗体覆盖时自动回位的问题
		win32gui.SetForegroundWindow(hwnd)
	# 再次关闭2号屏幕手写屏,可能会和其他线程产生鼠标键盘的争用
	app.window(handle = hwnd).click_input(coords=(13, 13))
	app.window(handle = hwnd).type_keys('R')
	app.window(handle = hwnd).type_keys('{ENTER}')

	# 解锁
	lock.release()

# 设置窗体坐标
def setWndCoords(hwnd, show_type):
	if show_type == 'shuffle':
		fullX = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
		fullY = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
		scrX = fullX / 5 * 2
		scrY = fullY / 5 * 2
		coordX = fullX / 5 * 3
		coordY = fullY / 5 * 3
		win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, int(coordX), int(coordY), int(scrX), int(scrY),
							  win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER | win32con.SWP_SHOWWINDOW)

# 将控制台挪到轮播窗体后面
setWndCoords(app[u'Radmin Viewer'].handle, 'shuffle')

# 分屏展现17地市监控界面
for key, value in cities.items() :
	t = threading.Thread(target = connect, args = (key, value, 'split'))
	t.start()
	while t.is_alive():
		time.sleep(1)

# 轮播展现17地市监控界面
while True:
	for key, value in cities.items() :
		t = threading.Thread(target = connect, args = (key, value, 'shuffle'))
		t.start()
		while t.is_alive():
			time.sleep(1)
		time.sleep(15)