# coding: utf-8


class FuturePolicies(list):
	def __init__(self):
		super(FuturePolicies, self).__init__(self)

	def append(self, f_instance):
		if not isinstance(f_instance, FutureZenPolicyItemModel):
			raise TypeError
		list.append(self, f_instance)


class FutureZenPolicyItemModel(object):
	def __init__(self):
		self._order_number = 0
		self._policy_id = 0
		self._source = []
		self._destination = []
		self._service = []

		self._action = "Allow"
		self._schedule = ""
		self._QoS_limit_session = ""
		self._session_limit = ""
		self._timeout = 600
		self._log = 0
		self._ddos = ""

		self._applications = "", # 애플리케이션
		self._http_filtering = "", #	HTTP 필터링
		self._ips = "all", # IPS
		self._anti_virus = "", #	안티바이러스
		self._anti_spam = "", # 안티스팸

		self._cross_policy = "off", # 양방향정책
		self._use = "on", # 동작
		self._policy_valid_time = "", # 정책유효기간
		self._date_time = "", #시간
		self._memo = "", # 메모
		self._ssli = "off" # SSLI

	@property
	def order_number(self):
		return self._order_number
	@order_number.setter
	def order_number(self, v):
		self._order_number = v

	@property
	def policy_id(self):
		return self._policy_id
	@policy_id.setter
	def policy_id(self, v):
		self._policy_id = v

	@property
	def source(self):
		return self._source
	@source.setter
	def source(self, v):
		if v['source_object_type'] in ("Any", "interface", "address", "user", "domain", "country",):
			pass
		elif v['source_object_type'] in ("user_group", "addr_group",):
			if not v["source_object_group"]:
				raise ValueError
			for o in v["source_object_group"]:
				if not o:
					raise ValueError
		else:
			raise ValueError


		self._source.append(v)

	@property
	def destination(self):
		return self._destination
	@destination.setter
	def destination(self, v):
		if v['destination_object_type'] in ("Any", "interface", "address", "user", "domain", "country",):
			pass
		elif v['destination_object_type'] in ("user_group", "addr_group",):
			if not v["destination_object_group"]:
				raise ValueError
			for o in v["destination_object_group"]:
				if not o:
					raise ValueError
		else:
			raise ValueError
		self._destination.append(v)

	@property
	def service(self):
		return self._service
	@service.setter
	def service(self, v):
		# v = {
		# 	"service_group": "Any", # 서비스그룹
		# 	"service_object_name": "Any", # 서비스객체이름
		# 	"protocol": "", # 프로토콜
		# 	"source_port": "", # 출발포트
		# 	"destination_port": "", # 목적포트
		# }
		self._service.append(v)

	@property
	def action(self):
		return self._action
	@action.setter
	def action(self, v):
		if v not in ("Allow", "Deny", "IPSec",):
			raise ValueError
		self._action = v

	@property
	def schedule(self):
		return self._schedule
	@schedule.setter
	def schedule(self, v):
		self._schedule = v

	@property
	def QoS_limit_session(self):
		return self._QoS_limit_session
	@QoS_limit_session.setter
	def QoS_limit_session(self, v):
		self._QoS_limit_session = v

	@property
	def session_limit(self):
		return self._session_limit
	@session_limit.setter
	def session_limit(self, v):
		try:
			v = int(v)
			if v < 0 or v > 1200:
				raise ValueError
		except TypeError:
			raise
		self._session_limit = v

	@property
	def timeout(self):
		return self._timeout
	@timeout.setter
	def timeout(self, v):
		try:
			v = int(v)
			if v < 0 or v > 1200:
				raise ValueError
		except TypeError:
			raise
		self._timeout = v

	@property
	def log(self):
		return self._log
	@log.setter
	def log(self, v):
		self._log = v

	@property
	def schedule(self):
		return self._schedule
	@schedule.setter
	def schedule(self, v):
		self._schedule = v
