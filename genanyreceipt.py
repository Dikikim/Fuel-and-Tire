import itertools
import random
import re
from datetime import datetime
from io import BytesIO
from typing import Dict, Iterable, Union, List, Callable

from PIL import Image

import bulk_charge
import fts_structs
import heartbeatstore
import minion_settings
from checktire import check_dot, check_sidewall, check_tread_depth, check_tread, check_punc
from fts_structs import Service, TireData, Routine
from globals import Maint, Data
from otistructs import AuthorizationDetails
from pdfgen import PDFGen

avg_gas_price = 3.00  # $
_abbr = {"forward": "Fwd", "drive": "Dr"}
_pattern_combine_spaces = re.compile(r" {2,}")  # replace 2 or more spaces with a single space
_pattern_simplify_abbr = re.compile(r"(?<=\b.) (?=.\b)")  # delete space between single letters


def create_receipt(service: Service, data: Dict[str, TireData], comment: str):
	if service.routine == Routine.REPAIR:
		return create_misc_receipt(service, comment)
	elif service.routine == Routine.AUDIT:
		return create_assessment_receipt(service, data, comment)
	else:  # INFLATION, PURGE_FILL, VERIFICATION
		if service.type_id == 'tireassurance':
			return create_tiir_receipt(service, data, comment)
		else:
			return create_tips_receipt(service, data, comment)


def create_declined_receipt():
	byte_buffer = BytesIO()  # used as a write-capable file in memory
	pdf = PDFGen(byte_buffer, 2.75, 0)  # 2.75 inches wide with 0 inch margins

	builder = PDFBuilder(pdf)  # service and tire data are only used for printing tire results

	builder.add_contact_info([])
	builder.add_authorization_details()
	builder.add_closing()

	pdf.finish()
	byte_buffer.seek(0)
	return byte_buffer


def create_misc_receipt(service: Service, comment: str):
	byte_buffer = BytesIO()  # used as a write-capable file in memory
	pdf = PDFGen(byte_buffer, 2.75, 0)  # 2.75 inches wide with 0 inch margins
	builder = PDFBuilder(pdf, service, {})

	builder.add_contact_info(("Mobile Nitrogen Tire", "Inspection & Inflation"))
	builder.add_authorization_details()
	builder.add_service_details("Tire Inspection & Pressure Service", "TIPS")
	builder.pdf.addtable("Ea. Price", f"${service.default_price:.2f}")
	builder.pdf.addtable("Quantity", f"{Data.Vehicle.Config.repair_quantity}")
	builder.add_comment(comment)
	builder.add_closing()

	pdf.finish()
	byte_buffer.seek(0)
	return byte_buffer


def create_tips_receipt(service: Service, data: Dict[str, TireData], comment: str):
	byte_buffer = BytesIO()  # used as a write-capable file in memory
	pdf = PDFGen(byte_buffer, 2.75, 0)  # 2.75 inches wide with 0 inch margins
	builder = PDFBuilder(pdf, service, data)

	builder.add_contact_info(("Mobile Nitrogen Tire", "Inspection & Inflation"))
	builder.add_authorization_details()
	builder.add_service_details("Tire Inspection & Pressure Service", service.type + " TIPS")
	builder.add_tire_data_auto(service.image, service.image_scale)
	if "Steer" in service.fullname():  # FIXME: hot fix for this ONE service
		builder.add_savings_report(comment, no_savings=True)
	else:
		builder.add_savings_report(comment)
	builder.add_closing()

	pdf.finish()
	byte_buffer.seek(0)
	return byte_buffer


def create_assessment_receipt(service: Service, data: Dict[str, TireData], comment: str):
	byte_buffer = BytesIO()  # used as a write-capable file in memory
	pdf = PDFGen(byte_buffer, 2.75, 0)  # 2.75 inches wide with 0 inch margins
	builder = PDFBuilder(pdf, service, data)

	builder.add_contact_info(("Mobile Nitrogen Tire", "Inspection & Inflation"))
	builder.add_authorization_details()  # Free
	builder.add_service_details("Tire Assessment Service", service.type + " TAS")  # XXX: just guessing
	builder.add_tire_data_auto(service.image, service.image_scale, insp_rule=1, inflation=False, use_mrsp=True)
	builder.add_savings_report(comment, no_savings=True, no_nitrogen=True)
	builder.add_closing()

	pdf.finish()
	byte_buffer.seek(0)
	return byte_buffer


def create_tiir_receipt(service: Service, data: Dict[str, TireData], comment: str):
	byte_buffer = BytesIO()  # used as a write-capable file in memory
	pdf = PDFGen(byte_buffer, 2.75, 0)  # 2.75 inches wide with 0 inch margins
	builder = PDFBuilder(pdf, service, data)

	builder.add_contact_info(("Astrae Tire Assurance", "TIIR Service"))
	builder.add_authorization_details()
	if Data.Vehicle.company:
		builder.add_service_details("Tire Insp., Inflation & Replacement",
									Data.Vehicle.company + " - TIIR Service")
	else:
		builder.add_service_details("Tire Insp., Inflation & Replacement", "TIIR Service")
	builder.add_tire_data_auto(service.image, service.image_scale)
	builder.add_savings_report(comment, no_savings=True)
	builder.add_closing(add_tire_costs=True)

	pdf.finish()
	byte_buffer.seek(0)
	return byte_buffer


def create_bulk_receipt(charges: List[bulk_charge.ChargeType]):
	byte_buffer = BytesIO()  # used as a write-capable file in memory
	pdf = PDFGen(byte_buffer, 2.75, 0)  # 2.75 inches wide with 0 inch margins
	builder = PDFBuilder(pdf)

	builder.add_contact_info(("Mobile Nitrogen Tire", "Inspection & Inflation"))
	builder.add_authorization_details()
	builder.add_service_details("Tire Inspection & Pressure Service", "TIPS")

	pdf.options(10)
	for charge in charges:
		pdf.addtableadv(f"{charge['service']}", f"${charge['amount']:.2f}", True, True)
		pdf.addline(f"  {charge['ctrlnum']:010} - {charge['vehicle']}")

	builder.add_closing()

	pdf.finish()
	byte_buffer.seek(0)
	return byte_buffer


class PDFBuilder:
	def __init__(self, pdf: PDFGen, service: Service=None, data: Dict[str, TireData]=None):
		self.pdf = pdf
		self.service = service
		self.data = data
		self.now = datetime.now()

	def add_contact_info(self, service_description: Iterable[str]):
		self.pdf.addtable(Data.Times.accept.strftime("%m/%d/%Y"),
						  Data.Times.accept.strftime("%H:%M:%S"))
		self.pdf.insertimage("files/FTS Logo.jpg", scale=0.21, y_pad=2)  # add logo - adjust to about 66 pixels in height
		self.pdf.skip(70)
		self.pdf.options(9, True, align='center')
		self.pdf.addline("Fuel & Tire Saver Systems Company, LLC")
		self.pdf.addline("45915 Maries Rd, Suite 136")
		self.pdf.addline("Dulles, VA 20166")
		self.pdf.addline("703-429-0382")
		self.pdf.skip(5)
		self.pdf.options(14, True, True, 'center')
		for line in service_description:
			self.pdf.addline(line)
		self.pdf.skip(5)
		self.pdf.options(11, align='center')
		self.pdf.addline("www.fuelandtiresaver.com")

	def add_authorization_details(self):
		self.pdf.skip(5)
		trans = Data.Payment
		if trans.prepaid_code:
			self.pdf.options(14, align='center')
			if trans.prepaid_index is not None:
				# sometimes I use purposely invalid prepaid codes which aren't "prepaid" but rather a reason why there was no charge
				self.pdf.addline("Prepaid Code:")
			self.pdf.addline(trans.prepaid_code)
			self.pdf.skip(5)
			self.pdf.options(9, align='center')
			self.pdf.addtable("Value", "USD ${:.2f}".format(Data.Vehicle.Config.price))  # print service value amount for prepaids, including bulk
		elif trans.use_alt_billing():
			self.pdf.options(30, bold=True, align='center')  # HUGE MONEY
			self.pdf.addline(f"${trans.price_paid:.2f}")  # mike says so they know it's not free
			self.pdf.skip(5)
			self.pdf.options(14, align='center')
			if Data.Payment.is_voided():  # check for void not ok because nobody was literally charged and is_ok will return false
				self.pdf.addline("Transaction VOIDED")
			else:
				self.pdf.addline("Charged to FTS Internal")
			# alt account contains 5 digits. billing contains short string
			self.pdf.addline(f"Acct. {trans.alt_account} {trans.alt_billing.value or ''}")
			self.pdf.skip(5)
		else:
			self.pdf.options(9, align='center')
			try:
				rdict = trans.rdict
				if isinstance(rdict, dict):  # COF
					if "transaction_type" in rdict and rdict["transaction_type"] == "charge":  # COF transaction
						self.pdf.addtable("Transaction ID", str(rdict["host_transaction_id"]))
						self.pdf.addtable("MID", (len(str(rdict["mid"])) - 7) * "*" + str(rdict["mid"])[-7:])
						self.pdf.addtable("TID", str(rdict["tid"]))

						self.pdf.addline("Purchase")
						self.pdf.addtable('PAN', "************" + str(rdict["tokenized_card_info"]["last_four"]))
						self.pdf.addtable("Method", "Card On File")
						self.pdf.addtable("Auth Mode", "Issuer")
						self.pdf.addtable("Amount", "USD ${:.2f}".format(trans.price_paid))
						self.pdf.addtable("Auth Code", str(rdict["auth_code"]))
					else:
						self.pdf.addline("Unknown transaction")
				elif isinstance(rdict, AuthorizationDetails):  # reader
					self.pdf.addtable("Transaction ID", rdict.transaction_db_id)
					self.pdf.addtable("MID", "***1098869")
					self.pdf.addtable("TID", rdict.tid)

					self.pdf.addline("Purchase")
					if rdict.partial_pan != '':
						self.pdf.addtable(rdict.card_type, "************" + rdict.partial_pan)
					else:
						self.pdf.addline("No card")  # there was an error reading the card
					self.pdf.addtable("Method", "Chip" if rdict.channel == "Contact" else rdict.channel)
					self.pdf.addtable("CVM", rdict.cvm)
					self.pdf.addtable("Auth Mode", "Issuer")
					self.pdf.addtable("Amount", "USD ${:.2f}".format(trans.price_paid))
					self.pdf.addtable("Auth Code", rdict.auth_id)

					self.pdf.addline("EMV Details")
					self.pdf.addtable("AID", rdict.aid)
					self.pdf.addtable("TVR", rdict.tvr)
					self.pdf.options(8)
					self.pdf.addtable("IAD", rdict.iad)
					self.pdf.options(9)
					self.pdf.addtable("TSI", rdict.tsi)
					self.pdf.addtable("ARC", rdict.arc)
			except KeyError:
				self.pdf.addtable("AP CODE", "None")  # Credit card failed

	def add_service_details(self, description: str, service_details: str):
		self.pdf.skip(2)
		self.pdf.options(9)
		self.pdf.addtable("Description:", description)
		self.pdf.addtable("Email:", Data.Contact.email or "<No Email>")
		if Data.Vehicle.mileage is not None:
			self.pdf.addtable("Mileage:", f"{Data.Vehicle.mileage:,}")  # print with thousands separators
		if Data.Payment.check_number:
			self.pdf.addtable("Check #:", f"{Data.Payment.check_number}")
		if Data.Vehicle.address:
			address1, address2 = Data.Vehicle.address.split("\n")
			if address1:
				self.pdf.addtable("Street:", f"{address1}")
			if address2:
				self.pdf.addtable("City:", f"{address2}")
		self.pdf.addtable("Control Number:", "{:010}".format(Data.control_number_as_int()))
		self.pdf.addtable("Miosk ID:", Maint.mioskid)
		self.pdf.skip(5)
		self.pdf.options(12, True, align='center')
		self.pdf.addline(service_details)
		if self.service:
			self.pdf.skip(5)
			self.pdf.options(18, True, align='center')
			self.pdf.addline(self.service.shortname())
		self.pdf.skip(5)
		# one or more of these will trigger
		if Data.Vehicle.plate_number:
			self.pdf.options(16, True, align='center')
			self.pdf.addline(f"{Data.Vehicle.plate_state or ''} {Data.Vehicle.plate_number}")
		if Data.Vehicle.vehicle_number:
			self.pdf.options(16, align='center')
			self.pdf.addline("#" + Data.Vehicle.vehicle_number)
		if Data.Vehicle.vin:
			self.pdf.options(12, align='center')
			self.pdf.addline(Data.Vehicle.vin)
		self.pdf.skip(10)

	def add_tire_data_auto(self, image: Union[BytesIO, str] = None, imgscale=1.0, insp_rule=2, inflation=True, use_mrsp=False):
		"""
		Automatically adds tire data and image.
		Abbreviates names for tires by slowly cutting off parts of the name until it can fit within a few characters.

		:param image: vehicle image to display in the center
		:param imgscale: image scale to have it fit, usually determined by the service class itself
		:param insp_rule: inpsection rule for printing info. integer from 0-2.<br/>
				<ol start=0>
				<li>prints no inspection info</li>
				<li>prints inspection data for tires that have a tread depth or DOT</li>
				<li>verbose. prints all inspection info whether we have it or not</li>
				</ol>
		:param inflation: if True, prints Act instead of Before and omits After
		:param use_mrsp: if True, prints MRSP instead of SP
		"""

		def getabbr(word: str):
			lword = word.lower()
			if lword in _abbr:
				return _abbr[lword]
			return word[0].upper()  # first letter is abbreviation if neither matches

		def simplify(phrase: str):
			a = _pattern_combine_spaces.sub(' ', phrase)
			b = _pattern_simplify_abbr.sub('', a)
			return b.strip()

		def replword(string: str, replace: str, replacement: str = None):
			if replacement is None:
				replacement = getabbr(replace)
			return re.sub(r"\b" + re.escape(replace) + r"\b", replacement, string)

		# bike detect
		oneside = False
		if self.service.numtires == 2 and self.service.numaxles == 2:
			# rework double axle on one side to single axle with front on left and rear on right
			oneside = True
			fa = self.service.template.axles[0]
			ra = self.service.template.axles[1]
			ft = (fa.left or fa.right)[0].label
			rt = (ra.left or ra.right)[0].label
			custom_axles = [fts_structs.Template.Axle("", [ft], [rt], 0, [""])]  # pass in left/right for getting correct data
		else:
			custom_axles = self.service.template.axles  # axles are normal

		last_pix_y = self.pdf.pix_y
		for axle in custom_axles:
			onesideside = None
			# so we know which side and so it doesn't print left or right
			if len(axle.right) == 0:
				onesideside = "l"
				oneside = True
			elif len(axle.left) == 0:
				onesideside = "r"
				oneside = True

			for i, (lt, rt) in enumerate(itertools.zip_longest(axle.left, axle.right)):  # if right does not exist, the right side will be None
				# abbreviate tire title text
				innout = self.service.template.inout_labels[len(axle.left)][i]
				if oneside and onesideside is None:  # special sideways one-sided dual axle
					left_tire_label = f"{self.service.template.axles[0].title} Tire"
					right_tire_label = f"{self.service.template.axles[1].title} Tire"
				else:
					left_tire_label = "{} {} Tire".format(axle.title, innout)
					right_tire_label = "{} {} Tire".format(axle.title, innout)
				if not oneside:
					left_tire_label = "Left " + left_tire_label
					right_tire_label = "Right " + right_tire_label
				level = 0
				while len(left_tire_label) > 15:  # reduce tire text to at most 15 chars by abbreviating one step at a time
					if level == 0:
						# Left -> L, Right -> R
						left_tire_label = replword(left_tire_label, "Left")
						right_tire_label = replword(right_tire_label, "Right")
					elif level == 1:
						# first word of the axle title -> abbr
						axle_1stword = axle.title.split(" ")[0]
						left_tire_label = replword(left_tire_label, axle_1stword)
						right_tire_label = replword(right_tire_label, axle_1stword)
					elif level == 2:
						# Tire -> ''
						left_tire_label = replword(left_tire_label, "Tire", "")
						right_tire_label = replword(right_tire_label, "Tire", "")
					elif level == 3:
						# rest of axle title -> abbr
						for axle_word in axle.title.split(" ")[1:]:
							left_tire_label = replword(left_tire_label, axle_word)
							right_tire_label = replword(right_tire_label, axle_word)
					elif level == 4:
						# Inner -> I
						# Outer -> O
						# (others) -> abbr
						for label in innout:
							for lblword in label.split(" "):  # mostly inner/outer but accounts for many words in labels
								left_tire_label = replword(left_tire_label, lblword)
								right_tire_label = replword(right_tire_label, lblword)
					else:
						break  # level too high, forget it

					# cut out extraneous spaces
					left_tire_label = simplify(left_tire_label)
					right_tire_label = simplify(right_tire_label)

					level += 1

				left_tire_label = simplify(left_tire_label)
				right_tire_label = simplify(right_tire_label)

				# add the data depending on the verbosity
				ld = self.data[lt.label] if lt else None
				rd = self.data[rt.label] if rt else None
				if insp_rule == 0:  # least vebose == print no inspection info
					linsp = rinsp = False
				elif insp_rule == 1:  # less verbose == inspection info depends on the data
					linsp = bool(ld.tread_depth or ld.dot) if ld else False
					rinsp = bool(rd.tread_depth or rd.dot) if rd else False
				else:  # most verbose == print all inspection info
					linsp = rinsp = True
				if ld is None:
					left_tire_label = ""
				if rd is None:
					right_tire_label = ""
				self.add_tire_data(left_tire_label, right_tire_label, ld, rd, left_insp=linsp, right_insp=rinsp, inflation=inflation, use_mrsp=use_mrsp)
			oneside = False  # reset for next set of tires

		tire_info_pixs = last_pix_y - self.pdf.pix_y  # the number of pixels the tire info took up
		if image:
			# give half the amount of space to negative padding which moves the image up to align with the info
			self.pdf.insertimage(image, imgscale, y_align='center', y_pad=-tire_info_pixs // 2)

	def add_tire_data(self, left_title, right_title, l: TireData = None, r: TireData = None,
					  left_insp=True, right_insp=True, inflation=True, use_mrsp=False):
		"""
		Adds all tire data for left and right side given to the receipt

		:param left_title: left tire name
		:param right_title: right tire name
		:param l: left tire data
		:param r: right tire data
		:param left_insp: print inspection info for left tire
		:param right_insp: print inspection info for right tire
		:param inflation:  if True, prints Act instead of Before and omits After
		:param use_mrsp: if True, prints MRSP instead of SP
		"""
		self.pdf.options(10, True)
		self.pdf.addtable(left_title, right_title)
		self.pdf.skip(2)
		self.pdf.options(8)

		def add_data(display: str, data_func: Callable[[TireData], str], check_func: Union[Callable[[TireData], bool], bool], insp=False):
			# data is replaced with a call to data_func, bolding is determined by a call to check_func
			# if a key does not exist on only one side, it should not be printed for that side
			# if both keys do not exist (left and right), that line will be skipped
			# data is not added if the tire data does not exist or the data function returns None
			# data is also not added if minimal is true for that side (unless always is True)
			# insp fields are only printed if left_insp or right_insp are true, respectively

			left_str, left_bold = "", False
			right_str, right_bold = "", False

			if l and (not insp or left_insp):
				data = data_func(l)
				if data is not None:
					left_str, left_bold = f"{display}: {data}", check_func if isinstance(check_func, bool) else check_func(l)
			if r and (not insp or right_insp):
				data = data_func(r)
				if data is not None:
					right_str, right_bold = f"{data} :{display}", check_func if isinstance(check_func, bool) else check_func(r)

			if left_str or right_str:
				self.pdf.addtableadv(left_str, right_str, left_bold, right_bold)

		add_data("SW", lambda x: x.sidewall_str(), lambda x: check_sidewall(x.sidewall_str()), insp=True)
		add_data("T", lambda x: x.tread_str(), lambda x: check_tread(x.tread_str()), insp=True)
		add_data("TD", lambda x: x.tread_depth_32(), lambda x: check_tread_depth(x.tread_depth_float()), insp=True)
		add_data("Punc", lambda x: x.puncture_str(), lambda x: check_punc(x.puncture_str()), insp=True)
		add_data("DOT", lambda x: x.dot, lambda x: check_dot(x.dot), insp=True)
		if inflation:
			add_data("Temp", lambda x: x.temp_str(), False)
		add_data("Before" if inflation else "Act", lambda x: f"{x.accurate_uncorrected():.2f} PSI", False)
		if inflation:
			add_data("After", lambda x: f"{x.corrected:.2f} PSI", False)
		add_data("Diff", lambda x: f"{x.diff():.2f} PSI", True)
		add_data("N2", lambda x: f"{x.nitrogen:.1f}%" if x.nitrogen else None, True)
		add_data("MRSP" if use_mrsp else "SP", lambda x: f"{x.sp():.0f} PSI" if use_mrsp else f"{x.sp():.2f} PSI", True)

	def add_comment(self, comment: str):
		if comment:  # add comment if it's not null or empty
			self.pdf.skip(5)
			self.pdf.options(13, align='center')
			self.pdf.addlineadv("Service Note:")
			self.pdf.options(11)

			def textwidth(s):
				return self.pdf.textwidth(s, 11, italic=True)

			# basic word wrapping algorithm
			# this means we assume there are super long words (which would not fit on a single line)
			spaces = re.compile(" +")
			line_length = self.pdf.width * 0.95
			partial = ""
			comment = comment.replace('\r', '').replace('\n', ' \n ')  # if a \r\n was placed, normalize all to \n
			for word in spaces.split(comment):
				if textwidth(partial) + textwidth(word) > line_length or word == '\n':  # adding this word would be too long, or forced newline
					self.pdf.addlineadv(partial, italic=True)
					partial = ""
				if word != '\n':
					partial += word + " "
			if len(partial.strip()) > 0:  # if theres still a bit more after the sentence ends, add it
				self.pdf.addlineadv(partial, italic=True)
			self.pdf.options(12, align='center')
			self.pdf.skip(5)

	def add_savings_report(self, comment: str=None, no_savings=False, no_nitrogen=False):
		"""
		Add the entire savings report, computing all savings from the data previously provided

		:param comment: technician comment to display. should contain newline characters to print on multiple lines
		:param no_savings: if tires were merely measured and not inflated, set to True so savings are not printed
		:param no_nitrogen: if tires were not inflated, set to True so nitrogen message is not printed
		"""
		tiredatas = list(self.data.values())
		# savings calculations. if savings is printed, be strict with the difference
		strict = not no_savings
		difs = [tiredata.diff(strict) for tiredata in tiredatas if tiredata.valid(strict)]  # SP - before for each tire, remove null before pressures
		if len(difs) == 0:  # avoid divide by 0
			avg_under = 0
		else:
			avg_under = sum(difs) / len(difs)

		# updated calculations for vehicles with 6 tires or more
		if len(tiredatas) >= 6:
			# 6 tires or more. this is between 1800-2500 miles per year, at $2.50/gal (as of jan 2019), at 6.8 MPG
			saved_per_psiui = 0.82
			mi_yr_min = 1800
			mi_yr_max = 2500
			mi_gal_best = 6.8
			mi_gal_worst = 6.8
		else:
			# 4 tires or less. this is 30000 miles per year, at $2.50/gal (as of jan 2019), between 13-28 MPG
			saved_per_psiui = 0.3
			mi_yr_min = 30000
			mi_yr_max = 30000
			mi_gal_best = 28.0
			mi_gal_worst = 13.0

		savings_percent = max(-avg_under, 0.0) * saved_per_psiui  # save % for every psi tire is underinflated
		savings_decimal = savings_percent / 100
		gas_price = minion_settings.get_float(minion_settings.Keys.GAS_PRICE) or avg_gas_price
		savings_min = mi_yr_min / mi_gal_best * gas_price * savings_decimal  # (mi/yr) / (mi/gal) * ($/gal) == $/year
		savings_max = mi_yr_max / mi_gal_worst * gas_price * savings_decimal

		# calculate max UI
		if len(difs) == 0:  # avoid /0 and zero-length array min/max errors
			maxpui = 0
			maxui = 0
			avgpui = 0
		else:
			percent_uis = [(tiredata.diff(strict)) / tiredata.sp() for tiredata in tiredatas if tiredata.valid(strict) and tiredata.sp() != 0]
			uis = [tiredata.diff(strict) for tiredata in tiredatas if tiredata.valid(strict) and tiredata.sp() != 0]
			if len(percent_uis) > 0:
				maxpui = min(percent_uis)
				maxui = min(uis)
				avgpui = sum(percent_uis) / len(percent_uis)
			else:
				maxpui = 0
				maxui = 0
				avgpui = 0

		self.pdf.skip(10)
		self.pdf.options(12, align='center')
		self.pdf.addlineadv("Most Severe UI: {:.2f} PSI".format(maxui), True)
		self.pdf.addlineadv("%UI: {:.2f}%".format(100 * maxpui), True)
		self.pdf.addlineadv("Avg UI: {:.2f} PSI".format(avg_under), True)
		self.pdf.addlineadv("%Avg UI: {:.2f}%".format(100 * avgpui), True)
		self.pdf.skip(5)
		self.add_comment(comment)

		if not no_savings and avg_under != 0:
			if avg_under < 0:  # underinflated
				tire_savings_min = 0
				tire_savings_max = 0
				if "bus" in self.service.vehicle.lower():  # bus types get the generic text
					self.pdf.addline("Our service increases fuel economy")
					self.pdf.addline("by 3-10% and can reduce tire wear")
					self.pdf.addline("and extend life as much as 20%")
					self.pdf.addline("while greatly improving safety!")
				else:
					self.pdf.addline("By proper inflation,")
					self.pdf.addlineadv("you just saved {:.1f}%".format(savings_percent), True)
					self.pdf.addline("fuel economy.")
					self.pdf.options(10, align='center')
					self.pdf.skip(10)
					self.pdf.addline("Your average fuel savings:")
					self.pdf.addline("Contact our office: (703) 429-0382")
					# self.pdf.addlineadv("between ${}-${} annually".format(int(savings_min), int(savings_max)), True)
			else:  # overinflated
				self.pdf.options(9, align='center')
				if avg_under <= 2:
					self.pdf.addline("Your tires were over-pressured.")
					tire_savings_min = 14
					tire_savings_max = 27
				elif avg_under <= 8:
					self.pdf.addline("Your tires were moderately over-pressured.")
					tire_savings_min = 27
					tire_savings_max = 53
				else:
					self.pdf.addline("Your tires were significantly over-pressured.")
					tire_savings_min = 53
					tire_savings_max = 80
				self.pdf.addline("We have set them to the proper pressure.")
				self.pdf.skip(10)
				self.pdf.addline("Your estimated savings is:")
				self.pdf.options(10, align='center')
				self.pdf.addlineadv("between ${}-${} annually".format(tire_savings_min, tire_savings_max), True)
				self.pdf.options(9, align='center')
				self.pdf.addline("from excessive tire wear.")
			avg_dollars = (savings_min + savings_max + tire_savings_min + tire_savings_max) / 2.0  # either savings or tire_savings will be 0. so just divide by 2
			heartbeatstore.increment("recent_saved_dollars", avg_dollars)
			heartbeatstore.increment("total_saved_dollars", avg_dollars)

		if not no_nitrogen:
			self.pdf.skip(5)
			self.pdf.options(9, italic=True, align='center')
			self.pdf.addline("Tires inflated with " + str(Maint.vals["nitrogen_percent"]) + "% pure FTS")
			self.pdf.options(11, italic=True, bold=True, align='center')
			self.pdf.addline("ÅASTRAEA Nitrogen")

	def add_closing(self, add_tire_costs=False):
		self.pdf.skip(10)
		self.pdf.options(10, align='center')
		if Data.control_number is not None:
			self.pdf.insertbarcode("{:010}".format(Data.control_number), x_offset=64)  # offset to try to center control number bar code
		if Data.Vehicle.vin:  # don't print a barcode of nothing
			self.pdf.insertbarcode(Data.Vehicle.vin, x_offset=14)  # offset to try to center vin bar code
		if not Data.Payment.prepaid_code and not Data.Payment.use_alt_billing():
			if add_tire_costs:
				cfg = Data.Vehicle.Config
				self.pdf.addtable("TIIR Service Labor:", "${:.2f}".format(cfg.labor_cost))
				tirecost = cfg.tire_cost * cfg.tire_decimal
				self.pdf.addtable("Tire Cost:", "${:.2f}".format(tirecost))
				salestax = cfg.tire_cost * cfg.sales_tax_rate
				self.pdf.addtable("Sales Tax:", "${:.2f}".format(salestax))
				rfee = cfg.recycling_fee * cfg.tire_decimal
				self.pdf.addtable("Tire Recycling Fee - Qty {:.2f}:".format(cfg.tire_decimal), "${:.2f}".format(rfee))
				self.pdf.skip(3)
				self.pdf.drawline(0, 1)
				total_payment = cfg.labor_cost + tirecost + salestax + rfee
				self.pdf.addtable("Total Monthly Payment:", "${:.2f}".format(total_payment))
				self.pdf.skip(3)
				heartbeatstore.increment("recent_material_costs", tirecost)
				heartbeatstore.increment("total_material_costs", tirecost)
				heartbeatstore.increment("recent_salestax_collected", salestax)
				heartbeatstore.increment("total_salestax_collected", salestax)
				heartbeatstore.increment("recent_tire_recycling_fees", rfee)
				heartbeatstore.increment("total_tire_recycling_fees", rfee)
			self.pdf.addtable("Charged Amount:", "${:.2f}".format(Data.Payment.price_paid))
			self.pdf.addtable("Response:", str(Data.Payment.status))
		if Data.Payment.is_ok():  # dont print for declined receipt
			self.pdf.addline("No Refunds,")
			self.pdf.addline("Service Credit Only.")
			self.pdf.skip(5)
			self.pdf.addline("I agree to pay above total amount")
			self.pdf.addline("according to card issuer agreement.")
			self.pdf.addline("Retain this copy for your")
			self.pdf.addline("statement verification.")
		self.pdf.skip(3)
		self.pdf.addline("Cardholder Copy")
		self.pdf.skip(3)
		self.pdf.options(10)
		self.pdf.addline("Service End Time:")  # time receipt is printed
		self.pdf.addtable(self.now.strftime("%m/%d/%Y"), self.now.strftime("%H:%M:%S"))
		self.pdf.options(14, True, True, 'center')
		self.pdf.skip(5)
		self.pdf.addline("Thanks for your business!")
		self.pdf.skip(5)
		self.pdf.options(8, align='center')
		self.pdf.addline(chr(169) + "2023 Fuel & Tire Saver Sys Co, LLC")  # 169 is copyright (c)
		self.pdf.addline("All Rights Reserved.")
		if Maint.coupon_img:
			if Maint.coupon_img == "RANDOM":
				verified_code = random.choice(list(Maint.valid_coupon_codes.values()))  # choose random from all filenames in RMG
			else:
				verified_code = Maint.coupon_img
			self.pdf.skip(5)
			self.pdf.insertimage(verified_code, scale='fit')
			img = Image.open(verified_code)  # open it just so we can get the height and skip
			self.pdf.skip(self.pdf.width / img.width * img.height)
		self.pdf.skip(2)
