# Copyright (c) 2024, itsdave GmbH and contributors
# For license information, please see license.txt

import frappe
from frappe.model.document import Document
from iiq_check_connect.tools import prepare_export as tools_prepare_export

class iiQCheckFunctions(Document):
	@frappe.whitelist()
	def prepare_export(self):
		tools_prepare_export(interactive=True)
