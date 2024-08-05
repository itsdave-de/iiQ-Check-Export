import datetime
import frappe
import pandas as pd
import openpyxl
import io
from frappe.utils.file_manager import save_file
from frappe.core.api.file import create_new_folder
from frappe import _
from ftplib import FTP, FTP_TLS
import logging

def convert_frappe_dict_to_dict(data):
    return [dict(item) if isinstance(item, frappe._dict) else item for item in data]

@frappe.whitelist()
def prepare_export(interactive=False):
    settings = frappe.get_single("iiQ-Check Settings")
    if not interactive:
        if not settings.enable_job:
            message = _("iiQ-Check export is disabled, exiting.")
            print(message)
            if interactive: frappe.throw(message)
            return
    
    if not settings.einheit_kategorie:
        message = _("No Eingeheit-Kategorie for export selected.")
        print(message)
        if interactive: frappe.throw(message)
        return
    
    if not settings.kundentyp:
        message = _("No Kundentyp selected for export.")
        print(message)
        if interactive: frappe.throw(message)
        return

    # Number of days you want to subtract
    days_to_subtract = settings.export_days_after_departure

    # Current date and time
    current_date = datetime.datetime.now()

    # Calculated past date at the start of the day
    start_of_day = (current_date - datetime.timedelta(days=days_to_subtract)).replace(hour=0, minute=0, second=0, microsecond=0)

    # Check if an export for the same departure date already exists
    existing_export = frappe.get_all("iiQ-Check Export", filters={"departure_date": start_of_day.date()}, fields=["name"])
    if existing_export:
        message = _(f"An export for the departure date {start_of_day.date()} already exists. Aborting the operation.")
        print(message)
        if interactive: frappe.throw(message)
        return

    # End of that day
    end_of_day = start_of_day + datetime.timedelta(1)

    # Format the dates for SQL query
    start_date_str = start_of_day.strftime('%Y-%m-%d %H:%M:%S')
    end_date_str = end_of_day.strftime('%Y-%m-%d %H:%M:%S')

    # Construct IN clauses for einheit_kategorie and kundentyp
    einheit_kategorie_list = [ek.einheit_kategorie for ek in settings.einheit_kategorie]
    kundentyp_list = [kt.kundentyp for kt in settings.kundentyp]

    # Format the IN clauses for SQL
    einheit_kategorie_str = ', '.join(f"'{ek}'" for ek in einheit_kategorie_list)
    kundentyp_str = ', '.join(f"'{kt}'" for kt in kundentyp_list)

    query = f"""
        SELECT
            kd.nachname as name,
            kd.anrede AS salutation,
            kd.email as email,
            kd.land AS language
        FROM tabReservierung res
        LEFT JOIN `tabCamping Kunde` kd ON res.kundennummer = kd.name
        WHERE res.abreise >= '{start_date_str}' AND res.abreise < '{end_date_str}'
        AND res.kategorie IN ({einheit_kategorie_str})
        AND kd.kundentyp IN ({kundentyp_str})
        AND kd.email IS NOT NULL AND kd.email != ''
        """

    data = frappe.db.sql(query, as_dict=1)

    # Convert frappe._dict to regular dict
    data = convert_frappe_dict_to_dict(data)

    # Check if data is not empty and is in the correct format
    if not data:
        message = _("No data returned from the query. Seems like there are no departures for the configured filter, or the import from Compusoft is not working correctly.")
        print(message)
        if interactive: frappe.throw(message)
        return

    # Debug print to inspect data format
    message = f"Query returned {len(data)} records."
    print(message)
    if interactive: frappe.msgprint(message)

    try:
        # Convert the data to a DataFrame
        df = pd.DataFrame(data)

        # Extract language mapping and default language from settings
        default_language = settings.default_language
        language_mapping = {el.country_code: el.language_string for el in settings.language_mapping}

        # Define the function to map language values with fallback to default
        def map_language(language):
            return language_mapping.get(language, default_language)

        # Apply the mapping function to the "language" column
        df["language"] = df["language"].apply(map_language)

        # Add departure_at column using start_date_str
        df["departure_at"] = start_of_day.strftime('%Y-%m-%d')

        # Save the DataFrame to a BytesIO object
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        message = "Query executed successfully. Data saved to BytesIO object."
        print(message)
        if interactive: frappe.msgprint(message)

        # Generate statistics
        statistics = f"Total Records: {len(data)}\n"
        statistics += f"Export Date: {current_date.strftime('%Y-%m-%d %H:%M:%S')}\n"

        # Create a new "iiQ-Check Export" document
        new_doc = frappe.get_doc({
            "doctype": "iiQ-Check Export",
            "created_on": current_date,
            "departure_date": start_of_day.date(),
            "number_of_recipients": len(data),
            "status": "prepared",
            "statistics": statistics
        })
        new_doc.insert()
        message = f"New iiQ-Check Export document created: {new_doc.name}"
        print(message)
        if interactive: frappe.msgprint(message)

        # Create the "iiq-check" folder if it does not exist
        folder_name = "Home/iiq-check"
        if not frappe.db.exists("File", {"file_name": "iiq-check", "folder": "Home"}):
            create_new_folder("iiq-check", "Home")
            message = f"Folder '{folder_name}' created."
            print(message)
            if interactive: frappe.msgprint(message)

        # Create a filename with the date of departure
        filename = f"iiq-check-export-{start_of_day.date()}.xlsx"

        # Attach the generated XLSX file to the new document and save it in the "iiq-check" folder
        file_doc = save_file(
            fname=filename,
            content=output.getvalue(),
            dt="iiQ-Check Export",
            dn=new_doc.name,
            folder=folder_name,
            is_private=1
        )
        message = f"File {filename} attached to iiQ-Check Export document {new_doc.name} in folder {folder_name}"
        print(message)
        if interactive: frappe.msgprint(message)

        # Link the file to the xlsx_file field in the document
        new_doc.xlsx_file = file_doc.file_url
        new_doc.status = "exported"
        new_doc.save()

        # Link to the new export document
        export_link = frappe.utils.get_url_to_form("iiQ-Check Export", new_doc.name)
        message = f"Export completed successfully. You can access the export document <a href='{export_link}'>here</a>."
        print(message)
        if interactive: frappe.msgprint(message)
        return new_doc.name

    except frappe.DocumentModifiedError as e:
        # If the document was modified by another process, re-fetch and save again
        new_doc = frappe.get_doc("iiQ-Check Export", new_doc.name)
        try:
            file_doc = save_file(
                fname=filename,
                content=output.getvalue(),
                dt="iiQ-Check Export",
                dn=new_doc.name,
                folder=folder_name,
                is_private=1
            )
            message = f"File {filename} attached to iiQ-Check Export document {new_doc.name} in folder {folder_name} after retrying"
            print(message)
            if interactive: frappe.msgprint(message)

            # Link the file to the xlsx_file field in the document
            new_doc.xlsx_file = file_doc.file_url
            new_doc.status = "exported"
            new_doc.save()

            # Link to the new export document
            export_link = frappe.utils.get_url_to_form("iiQ-Check Export", new_doc.name)
            message = f"Export completed successfully. You can access the export document <a href='{export_link}'>here</a>."
            print(message)
            if interactive: frappe.msgprint(message)

        except Exception as inner_e:
            message = f"Retry failed: {inner_e}"
            print(message)
            if interactive: frappe.throw(message)
            new_doc.status = "failed"
            new_doc.save()
            raise inner_e

    except Exception as e:
        # Update the status to failed in case of an error
        if 'new_doc' in locals():
            new_doc.status = "failed"
            new_doc.save()
        message = f"Error converting data to DataFrame, saving to Excel, or attaching file: {e}"
        print(message)
        if interactive: frappe.throw(message)
        raise e


@frappe.whitelist()
def upload_to_ftp(export_name):
    settings = frappe.get_single("iiQ-Check Settings")
    export_doc = frappe.get_doc("iiQ-Check Export", export_name)

    # Validate FTP settings
    if not all([settings.ftp_server, settings.ftp_user, settings.ftp_password, settings.ftp_path]):
        message = _("FTP settings are not fully configured. Please check the iiQ-Check Settings.")
        print(message)
        frappe.throw(message)

    if not export_doc.xlsx_file:
        message = _("No XLSX file attached to this export.")
        print(message)
        frappe.throw(message)

    file_url = export_doc.xlsx_file
    file_doc = frappe.get_doc("File", {"file_url": file_url})

    if not file_doc:
        message = _("Attached file not found.")
        print(message)
        frappe.throw(message)

    file_content = file_doc.get_content()

    if not file_content:
        message = _("No content found in the attached file.")
        print(message)
        frappe.throw(message)

    ftp_server = settings.ftp_server
    ftp_user = settings.ftp_user
    ftp_password = settings.get_password("ftp_password")
    ftp_path = settings.ftp_path
    ftp_port = settings.ftp_port if settings.ftp_port else 21
    use_secure_ftp = settings.use_secure_ftp

    ftp_log = []

    class CustomFTP(FTP):
        def sendcmd(self, cmd):
            response = super().sendcmd(cmd)
            ftp_log.append(f"Command: {cmd}")
            ftp_log.append(f"Response: {response}")
            return response

    class CustomFTP_TLS(FTP_TLS):
        def sendcmd(self, cmd):
            response = super().sendcmd(cmd)
            ftp_log.append(f"Command: {cmd}")
            ftp_log.append(f"Response: {response}")
            return response

    try:
        if use_secure_ftp:
            ftp = CustomFTP_TLS()
        else:
            ftp = CustomFTP()

        ftp.connect(ftp_server, ftp_port)
        ftp_log.append(f"Connected to FTP server {ftp_server} on port {ftp_port}.")

        ftp.login(user=ftp_user, passwd=ftp_password)
        ftp_log.append(f"Logged in as {ftp_user}.")

        if use_secure_ftp:
            ftp.prot_p()  # Switch to secure data connection
            ftp_log.append("Switched to secure data connection.")

        ftp.cwd(ftp_path)
        ftp_log.append(f"Changed directory to {ftp_path}.")

        with io.BytesIO(file_content) as f:
            ftp.storbinary(f'STOR {file_doc.file_name}', f)
            ftp_log.append(f"Uploaded file {file_doc.file_name}.")

        ftp.quit()
        ftp_log.append("FTP session closed.")

        message = f"File {file_doc.file_name} uploaded to FTP server {ftp_server}."
        print(message)
        frappe.msgprint(message)

        # Log the transfer in the export document's statistics
        current_date = datetime.datetime.now()

    except Exception as e:
        message = f"FTP upload failed: {e}"
        print(message)
        ftp_log.append(f"FTP upload failed: {e}")
        frappe.throw(message)
        raise e
    finally:
        # Ensure the logs are saved even if an error occurs
        current_date = datetime.datetime.now()
        formatted_log = f"""
            <div>
                <h4>FTP Upload Log - {current_date.strftime('%Y-%m-%d %H:%M:%S')}</h4>
                <pre>{'<br>'.join(ftp_log)}</pre>
            </div>
        """
        export_doc.statistics += formatted_log
        export_doc.save()


def hourly_job():
    log_message = ""
    status = ""
    reference_doctype = None
    reference_name = None
    
    try:
        print("Running iiQ Check hourly job...")
        settings = frappe.get_single("iiQ-Check Settings")
        
        if not settings.enable_job:
            log_message = "Job is disabled. Nothing to do."
            print(log_message)
            status = ""
            log_activity(status, log_message, reference_doctype, reference_name)
            return

        current_hour = datetime.datetime.now().hour

        if settings.export_hour != current_hour:
            log_message = f"Current hour ({current_hour}) does not match export hour ({settings.export_hour}). Skipping export."
            print(log_message)
            status = ""
            log_activity(status, log_message, reference_doctype, reference_name)
            return

        print("Job is enabled and the hour matches, starting export.")
        export_doc_name = prepare_export()

        if export_doc_name is None:
            log_message = "An export for the departure date already exists. Aborting the operation."
            print(log_message)
            status = "Failed"
        else:
            reference_doctype = "iiQ-Check Export"
            reference_name = export_doc_name
            frappe.db.commit()  # Commit after export preparation if it involves DB changes

            if settings.enable_ftp_export:
                upload_to_ftp(export_doc_name)
                log_message = f"Export successful. Export name: {export_doc_name}."
                status = "Sucess"
            else:
                log_message = "FTP export is disabled."
                status = "Sucess"
        
    except Exception as e:
        log_message = f"An error occurred: {str(e)}"
        print(log_message)
        status = "Failed"
        frappe.log_error(log_message, "iiQ Check hourly job error")

    finally:
        log_activity(status, log_message, reference_doctype, reference_name)
        frappe.db.commit()  # Commit the activity log entry

def log_activity(status, message, reference_doctype, reference_name):
    try:
        doc = frappe.get_doc({
            "doctype": "Activity Log",
            "subject": "iiQ-Check hourly job",
            "status": status,
            "content": message,
            "timestamp": datetime.datetime.now(),
            "reference_doctype": reference_doctype,
            "reference_name": reference_name
        })
        doc.insert()
    except Exception as e:
        frappe.log_error(f"Failed to log activity: {str(e)}", "Activity Log error")
