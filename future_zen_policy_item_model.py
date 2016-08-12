# coding: utf-8

# class ItemModel(FutureZenPolicyItemModel):
# 	def __init__(self):
# 		FutureZenPolicyItemModel.__init__(self)
#
# class SecrureWorksItemModel(ItemModel):
# 	def __init__(self):
# 		ItemModel.__init__(self)


from xlsx_handler import XlsxHandler
from DataStructure.item_future_zen_policy import FuturePolicies
from DataStructure.item_future_zen_policy import FutureZenPolicyItemModel
from DataStructure.return_value import ReturnValue
from DataStructure.matrix import Matrix



class ZenXlsxMaker(XlsxHandler):
	def __init__(self, model):
		if not isinstance(model, FuturePolicies):
			raise TypeError
		XlsxHandler.__init__(self, model)

		self.model = model

		self.subtitle = (
			'순위', '정책ID', '출발지 객체종류','출발지 객체그룹', '출발지 객체이름', '출발지 객체내용','목적지 객체종류','목적지 객체그룹', '목적지 객체이름', '목적지 객체내용',
			'서비스그룹', '서비스객체이름', '프로토콜', '출발포트', '목적포트', '행위', '스케쥴',
			'QoS', '세션제한', '타임아웃', '로그', 'DDoS', '애플리케이션', 'HTTP 필터링', 'IPS',
			'안티바이러스', '안티스팸', '양방향정책', '동작', '정책유효기간', '시간', '메모', 'SSLI')

		self._source_order_list = ("source_object_type","source_object_group", "source_object_name", "source_object_data",)


	def init_subtitle(self):
		"""최상단 각 컬럼 필드 이름 삽입"""
		for col, value in enumerate(self.subtitle, 0):
			self.matrix.set(0, col, value)
		self.current_row += 1
		return ReturnValue(True, None)


	def _insert_item(self, item):
		"""
		아이템을 Sheet에 기록한다.
		상대 좌표를 이용하여 기록
		"""
		m = Matrix()
		row = 0
		col = 0
		end_of_row = 0
		order = ("order_number", "policy_id", "source", "destination", "service", "action", "schedule", "QoS_limit_session",
		         "session_limit","timeout", "log", "ddos", "applications", "http_filtering", "ips", "anti_virus", "anti_spam", "cross_policy",
		         "use", "policy_valid_time", "date_time", "memo", "ssli",)

		for o in order:
			value = getattr(item, o)
			if o in ("source", "destination",):
				rtv = self._target(o, value)
				m.set_matrix(row, col, rtv["retval"]["matrix"])
				col += rtv["retval"]["col"]

				end_of_row = rtv["retval"]["row"] if rtv["retval"]["row"] > end_of_row else end_of_row
			elif o == "service":
				rtv = self._service(value)
				m.set_matrix(row, col, rtv["retval"]["matrix"])
				col += rtv["retval"]["col"]
				end_of_row = rtv["retval"]["row"] if rtv["retval"]["row"] > end_of_row else end_of_row

			else:
				m.set(row, col, value)
				col += 1
		return ReturnValue(True, {"matrix":m, "col": col, "row": end_of_row})


	def _target(self, target, data):
		m = Matrix()
		row = 0
		end_of_row = 0
		col = 0
		gi = 0
		# data = [
		# 	{
		# 		'source_object_type': 'Any',
		# 		'source_object_group': [
		# 			{'group_object_name': '', 'group_object_data': '', 'group_name': ''}
		# 		],
		# 		'source_object_data': '',
		# 		'source_object_name': 'Any'
		# 	}
		# ]
		for item in data:
			# 객체 종류
			m.set(row, col, item[target + "_object_type"])
			col += 1

			# 객체 그룹
			# 객체 종류가 Any라면 Group에 None. 아이템이 있다면
			if "group" in item[target + "_object_type"]:
				for i, obj_group in enumerate(item[target + "_object_group"]):
					m.set(row + end_of_row, col, obj_group["group_name"])
					for gi, d in enumerate(obj_group["group_object"], 1):
						group_m = Matrix()
						group_m.set(gi - 1, 0, d["name"])
						group_m.set(gi - 1, 1, d["data"])
						m.set_matrix(row + end_of_row, col + 1, group_m)
					end_of_row += gi
				col += 1
			else:
				m.set(row, col, None) # target + "_object_group"
				col += 1
				m.set(row, col, item[target + "_object_name"])
				col += 1
				m.set(row, col, item[target + "_object_data"])
				col += 1
		return ReturnValue(True, {"matrix":m, "col": 4, "row": end_of_row})


	def _service(self, data):
		m = Matrix()
		# {'service_group': '', 'source_port': '1~65535', 'destination_port': '1~65535', 'protocol': 'tcp', 'service_object_name': 'ALL(TCP)'}
		row = 0
		end_of_row = 0
		col = 0
		gi = 0

		for i, item in enumerate(data):
			group_m = Matrix()
			if item["service_group"]:
				group_m.set(row + i, col, item["service_group"])
			else:
				group_m.set(row + i, col, "")
			group_m.set(row + i, col + 1, item["service_object_name"]) # target + "_object_group"
			group_m.set(row+ i, col + 2, item["protocol"])
			group_m.set(row + i, col + 3, item["source_port"])
			group_m.set(row + i, col + 4, item["destination_port"])

			m.set_matrix(row + end_of_row, col, group_m)
			end_of_row += i

		return ReturnValue(True, {"matrix":m, "col": 5, "row": end_of_row})


if __name__ == "__main__":
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

	handler = ZenXlsxMaker(policies)
	handler.change_sheet_name(0, "WeGuardia_IPv4Filtering")
	handler.convert()
	handler.save()



# WeGuardia_IPv4Filtering