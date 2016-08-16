# coding: utf-8
import unittest

from future_zen_policy_item_model import ZenXlsxMaker
from future_zen_policy_item_model import FuturePolicies
from future_zen_policy_item_model import FutureZenPolicyItemModel

from xlsx_handler import openpyxl_handler

class ZenXlszMakerValueValidTest(unittest.TestCase):
	policies = FuturePolicies()
	c = FutureZenPolicyItemModel()

	data_ = {
		"source": [
			{
				"source_object_type": "Any", # 	출발지 객체종류
				"source_object_group": [
					{
						"group_name": "",
						"group_object": [],
					}
				], # 출발지 객체그룹
				"source_object_name": "Any", # 출발지 객체이름
				"source_object_data": "",# 출발지 객체내용
			}
		],
		"destination": [
			{
				"destination_object_type": "Any", # 	출발지 객체종류
				"destination_object_group": [
					{
						"group_name": "",
						"group_object_name": "",
						"group_object_data": ""
					}
				], # 출발지 객체그룹
				"destination_object_name": "Any", # 출발지 객체이름
				"destination_object_data": "",# 출발지 객체내용
			}
		],

		"service" : [
			{
				"service_group": "", # 서비스그룹
				"service_object_name": "Any", # 서비스객체이름
				"protocol": "", # 프로토콜
				"source_port": "", # 출발포트
				"destination_port": "", # 목적포트
			}
		],
	}
	c.order_number = 0
	c.policy_id = 1
	c.action = "Allow"
	c.schedule = "1"
	c.QoS_limit_session = "2"
	c.session_limit = 300
	c.timeout = 600
	c.log = 0
	c.ddos = "3"
	c.applications = "4" # 애플리케이션
	c.http_filtering = "" #	HTTP 필터링
	c.ips = "all" # IPS
	c.anti_virus = "" #	안티바이러스
	c.anti_spam = "" # 안티스팸
	c.cross_policy = "off" # 양방향정책
	c.use = "on" # 동작
	c.policy_valid_time = "" # 정책유효기간
	c.date_time = "" #시간
	c.memo = "" # 메모
	c.ssli = "off"

	for d in data_['source']:
		c.source = d
	for d in data_['destination']:
		c.destination = d
	for d in data_['service']:
		c.service = d

	policies.append(c)

	c = FutureZenPolicyItemModel()
	data_ = {
		"source": [
			{
				"source_object_type": "interface", # 	출발지 객체종류
				"source_object_group": [
					{
						"group_name": "",
						"group_object": [],
					}
				], # 출발지 객체그룹
				"source_object_name": "$eth2", # 출발지 객체이름
				"source_object_data": "",# 출발지 객체내용
			}
		],
		"destination": [
			{
				"destination_object_type": "Any", # 	출발지 객체종류
				"destination_object_group": [
					{
						"group_name": "",
						"group_object_name": "",
						"group_object_data": ""
					}
				], # 출발지 객체그룹
				"destination_object_name": "Any", # 출발지 객체이름
				"destination_object_data": "",# 출발지 객체내용
			}
		],

		"service" : [
			{
				"service_group": "service_group", # 서비스그룹
				"service_object_name": "AH", # 서비스객체이름
				"protocol": "AH", # 프로토콜
				"source_port": "", # 출발포트
				"destination_port": "", # 목적포트
			},
			{
				"service_group": "service_group", # 서비스그룹
				"service_object_name": "ALL(UDP)", # 서비스객체이름
				"protocol": "udp", # 프로토콜
				"source_port": "1~65535", # 출발포트
				"destination_port": "1~65535", # 목적포트
			}
		],
	}
	c.order_number = 1
	c.policy_id = 1
	c.action = "Allow"
	c.schedule = "1"
	c.QoS_limit_session = "2"
	c.session_limit = 300
	c.timeout = 600
	c.log = 0
	c.ddos = "3"
	c.applications = "4" # 애플리케이션
	c.http_filtering = "" #	HTTP 필터링
	c.ips = "all" # IPS
	c.anti_virus = "" #	안티바이러스
	c.anti_spam = "" # 안티스팸
	c.cross_policy = "off" # 양방향정책
	c.use = "on" # 동작
	c.policy_valid_time = "" # 정책유효기간
	c.date_time = "" #시간
	c.memo = "" # 메모
	c.ssli = "off"


	for d in data_['source']:
		c.source = d
	for d in data_['destination']:
		c.destination = d
	for d in data_['service']:
		c.service = d

	policies.append(c)

	c = FutureZenPolicyItemModel()
	data_ = {
		"source": [
			{
				"source_object_type": "addr_group", # 	출발지 객체종류
				"source_object_group": [
					{
						"group_name": "dd",
						"group_object": [
							{"name":"192.168.1.x", "data":"144.233.233.222;233.22.233.23;44.33.44.33"},
							{"name":"192.168.0.x", "data":"2.23;44.33.44.33"},
							{"name":"192.168.4.x", "data":"2.23;44.33.44.33"}
						]
					},
					{
						"group_name": "ip4_group",
						"group_object": [
							{"name":"1234", "data":"44.33.3.3"},
							{"name":"4567", "data":"33.3.3.3"},
							{"name":"467", "data":"2.2.2.2"},
							{"name":"457", "data":"1.1.1.1"}
						]
					},
				], # 출발지 객체그룹
				"source_object_name": "$eth2", # 출발지 객체이름
				"source_object_data": "",# 출발지 객체내용
			}
		],
		"destination": [
			{
				"destination_object_type": "address", # 	출발지 객체종류
				"destination_object_group": [
					{
						"group_name": "",
						"group_object_name": "",
						"group_object_data": ""
					}
				], # 출발지 객체그룹
				"destination_object_name": "192.168.1.x", # 출발지 객체이름
				"destination_object_data": "192.168.1.0/24",# 출발지 객체내용
			}
		],

		"service" : [
			{
				"service_group": "", # 서비스그룹
				"service_object_name": "ALL(TCP)", # 서비스객체이름
				"protocol": "tcp", # 프로토콜
				"source_port": "1~65535", # 출발포트
				"destination_port": "1~65535", # 목적포트
			}
		],
	}
	c.order_number = 1
	c.policy_id = 1
	c.action = "Allow"
	c.schedule = "1"
	c.QoS_limit_session = "2"
	c.session_limit = 300
	c.timeout = 600
	c.log = 0
	c.ddos = "3"
	c.applications = "4" # 애플리케이션
	c.http_filtering = "" #	HTTP 필터링
	c.ips = "all" # IPS
	c.anti_virus = "" #	안티바이러스
	c.anti_spam = "" # 안티스팸
	c.cross_policy = "off" # 양방향정책
	c.use = "on" # 동작
	c.policy_valid_time = "" # 정책유효기간
	c.date_time = "" #시간
	c.memo = "" # 메모
	c.ssli = "off"

	for d in data_['source']:
		c.source = d
	for d in data_['destination']:
		c.destination = d
	for d in data_['service']:
		c.service = d
	policies.append(c)


	handler = ZenXlsxMaker(policies, "zen_policies.xlsx")
	handler.new_workbook()
	handler.set_sheet_name("WeGuardia_IPv4Filtering")
	handler.convert()
	handler.save()

	sheet_handler = openpyxl_handler(filename="zen_policies.xlsx")
	sheet_handler.load_workbook()

	def test_get_sheet_names(self):
		self.assertEqual(self.sheet_handler.get_sheet_names()['retval'], [u'WeGuardia_IPv4Filtering', u'Sheet'])

	def test_get_row_number(self):
		self.assertEqual(self.sheet_handler.get_row_number()['retval'], 11)

	def test_get_col_number(self):
		self.assertEqual(self.sheet_handler.get_col_number()['retval'], 33)

	def test_get_row_values(self):
		t = []
		data = [[u'\uc21c\uc704', u'\uc815\ucc45ID', u'\ucd9c\ubc1c\uc9c0 \uac1d\uccb4\uc885\ub958', u'\ucd9c\ubc1c\uc9c0 \uac1d\uccb4\uadf8\ub8f9', u'\ucd9c\ubc1c\uc9c0 \uac1d\uccb4\uc774\ub984', u'\ucd9c\ubc1c\uc9c0 \uac1d\uccb4\ub0b4\uc6a9', u'\ubaa9\uc801\uc9c0 \uac1d\uccb4\uc885\ub958', u'\ubaa9\uc801\uc9c0 \uac1d\uccb4\uadf8\ub8f9', u'\ubaa9\uc801\uc9c0 \uac1d\uccb4\uc774\ub984', u'\ubaa9\uc801\uc9c0 \uac1d\uccb4\ub0b4\uc6a9', u'\uc11c\ube44\uc2a4\uadf8\ub8f9', u'\uc11c\ube44\uc2a4\uac1d\uccb4\uc774\ub984', u'\ud504\ub85c\ud1a0\ucf5c', u'\ucd9c\ubc1c\ud3ec\ud2b8', u'\ubaa9\uc801\ud3ec\ud2b8', u'\ud589\uc704', u'\uc2a4\ucf00\uc974', u'QoS', u'\uc138\uc158\uc81c\ud55c', u'\ud0c0\uc784\uc544\uc6c3', u'\ub85c\uadf8', u'DDoS', u'\uc560\ud50c\ub9ac\ucf00\uc774\uc158', u'HTTP \ud544\ud130\ub9c1', u'IPS', u'\uc548\ud2f0\ubc14\uc774\ub7ec\uc2a4', u'\uc548\ud2f0\uc2a4\ud338', u'\uc591\ubc29\ud5a5\uc815\ucc45', u'\ub3d9\uc791', u'\uc815\ucc45\uc720\ud6a8\uae30\uac04', u'\uc2dc\uac04', u'\uba54\ubaa8', u'SSLI'], [0L, 1L, u'Any', None, u'Any', None, u'Any', None, u'Any', None, None, u'Any', None, None, None, u'Allow', u'1', u'2', 300L, 600L, 0L, u'3', u'4', None, u'all', None, None, u'off', u'on', None, None, None, u'off'], [1L, 1L, u'interface', None, u'$eth2', None, u'Any', None, u'Any', None, u'service_group', u'AH', u'AH', None, None, u'Allow', u'1', u'2', 300L, 600L, 0L, u'3', u'4', None, u'all', None, None, u'off', u'on', None, None, None, u'off'], [None, None, None, None, None, None, None, None, None, None, u'service_group', u'ALL(UDP)', u'udp', u'1~65535', u'1~65535', None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [1L, 1L, u'addr_group', u'dd', u'192.168.1.x', u'144.233.233.222;233.22.233.23;44.33.44.33', u'address', None, u'192.168.1.x', u'192.168.1.0/24', None, u'ALL(TCP)', u'tcp', u'1~65535', u'1~65535', u'Allow', u'1', u'2', 300L, 600L, 0L, u'3', u'4', None, u'all', None, None, u'off', u'on', None, None, None, u'off'], [None, None, None, None, u'192.168.0.x', u'2.23;44.33.44.33', None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [None, None, None, None, u'192.168.4.x', u'2.23;44.33.44.33', None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [None, None, None, u'ip4_group', u'1234', u'44.33.3.3', None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [None, None, None, None, u'4567', u'33.3.3.3', None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [None, None, None, None, u'467', u'2.2.2.2', None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None], [None, None, None, None, u'457', u'1.1.1.1', None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None]]
		for n in range(0,11):
			t.append(self.sheet_handler.get_row_values(n)['retval'])
		self.assertEqual(t, data)






