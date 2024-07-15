// Copyright (c) 2024, itsdave GmbH and contributors
// For license information, please see license.txt

frappe.ui.form.on('iiQ-Check Export', {
	refresh: function(frm) {
		if (frm.doc.xlsx_file) {
			frm.add_custom_button(__('Upload to FTP'), function() {
				frappe.call({
					method: 'iiq_check_connect.tools.upload_to_ftp',
					args: {
						export_name: frm.doc.name
					},
					freeze: true,
					freeze_message: __('Uploading, please wait...'),
					callback: function(r) {
						if (r.message) {
							frappe.msgprint(r.message);
						}
						frm.reload_doc();  // Reload the form after the script has run
					}
				});
			});
		}
	}
});
