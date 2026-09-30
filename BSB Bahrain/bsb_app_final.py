"""
BSB Resource Manager v2
- Dark/Light mode with persistent preference
- Checkbox selection (Gmail-style) with bulk actions
- Excel export for selected items
- Settings panel in topbar
- BSB logo integration
- Maximized on startup
- Cormorant Garamond + Outfit fonts (from HTML v4)
"""
import customtkinter as ctk
from tkinter import messagebox, filedialog
import tkinter as tk
from tkinter import ttk
import pyodbc
from datetime import datetime
import os, json

# ── PERSISTENT SETTINGS ────────────────────────────────────
SETTINGS_FILE = os.path.join(os.path.expanduser('~'), '.bsb_settings.json')

def load_settings():
    try:
        with open(SETTINGS_FILE) as f:
            return json.load(f)
    except:
        return {'theme': 'dark'}

def save_settings(s):
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(s, f)
    except:
        pass

SETTINGS = load_settings()

# ── ROLE DEFINITIONS ───────────────────────────────────────
ADMIN_ROLES       = {'admin', 'procurement', 'it', 'operation'}
PROCUREMENT_ROLES = {'admin', 'procurement'}   # can access all procurement screens
STAFF_ROLES       = {'staff'}
ALL_ROLES         = ['admin','procurement','it','operation','staff','hr','fin']

# ── IT CATEGORY FIELD MAP ──────────────────────────────────
# Each IT category has its own set of fields (taken from the IT Excel).
# When adding/editing an asset, the user picks a category first, then
# these fields appear. Stored as JSON in ITAssets.DetailsJSON.
# The first field of each list is treated as the asset's display Name.
IT_CATEGORY_FIELDS = {
    'Active Servers':        ['Asset Name','Description','Roles','IP','Type of Server','Model','Specs','S/N','Status'],
    'Decomm Servers':        ['Asset Name','Description','Roles','IP','Type of Server','Model','Specs','S/N','Status'],
    'Network Devices':       ['Device Name','Model','Management IP','Location','Access Methods','IOS Version'],
    'IP Phones':             ['Ext. No','Name','MAC','SN','Asset Tag','Model','Brand','Remarks'],
    'CCTV':                  ['ID','Device Type','IPv4 Address','Software Version','IPv4 Gateway','Serial Number'],
    'CCTV Locations':        ['IDF No','Location','Cameras'],
    'CCTV Models':           ['Type','Model','Quantity'],
    'APs':                   ['Asset Tag','Model No.','Manufacturer','Serial','MAC Address'],
    'Printers':              ['Printer Name','Model','Location','IP','Asset Tag'],
    'UPS':                   ['Asset Name','Description','Roles','Model','S/N'],
    'iPads':                 ['Asset Tag','Model','Model No.','Serial','Default Location','School','Floor','Building','Status'],
    'VR':                    ['Asset Tag','Model','Serial No.','Location','Bundle SN','Intune','Remarks','Status'],
    'Mobile Devices':        ['Department','Device Model','Asset Tag','Mobile Number','Package'],
    'Chromebooks':           ['Asset Tag','Model','Model No.','Category','Manufacturer','Serial','Purchased','Supplier','Location','Warranty','School','Floor','Building'],
    'Macbooks & iMacs':      ['Asset Tag','Model No.','Manufacturer','Serial','Location','School','Floor','Building','RAM','CPU','HDD','Status'],
    'Desktops':              ['Asset Tag','Model','Serial Number','Class/Office','Floor','Building'],
    'Laptops':               ['Asset Tag','Model','Model No.','Category','Manufacturer','Serial','Default Location','Checked Out','Location'],
    'Projectors':            ['Brand','Model','Class/Office','Floor','Building'],
    'Cameras':               ['Asset Tag','Brand','Model','Class/Office','Floor','Building'],
    'Infants Camera':        ['Brand','Model','Serial Number'],
    'Smartboards':           ['Asset Tag','Brand','S/N','Model','Class/Office','Activation Key','Building'],
    'TV':                    ['Brand','S/N','Model','Class/Office'],
    'Audio Visual devices':  ['Model','Brand','Description','Location','Qty','Remarks'],
    'Access Control':        ['Location','Count'],
    'Kiosk Machines':        ['Asset Tag','Model','Serial Number','Device Status'],
    'Monitors n Docking Station':['Name','Position','School/Department','Assigned Device','Serial Number','Model'],
    # Decommission sub-mode categories
    'PC Laptop to Decom':    ['Type','Tag','Make/Model','Model','Serial','Location'],
    'iPads Macbooks to Decom':['Asset Tag','Model','Serial Number','Device Status'],
    'Smartboard to Decom':   ['Model','Brand','Serial Number','Location'],
}
IT_DEFAULT_FIELDS = ['Asset Tag','Model','Serial','Location']  # fallback for unknown categories

# Categories that belong in the "Decommission" sub-mode (not the main Manage grid)
IT_DECOM_CATEGORIES = ['PC Laptop to Decom','iPads Macbooks to Decom','Smartboard to Decom']

# Which field carries a broken/decommissioned status to colour the row?
# 'red'  values → fully broken/decommissioned (red row)
# 'amber' values → partially broken (yellow row)
IT_STATUS_FIELD = {
    'VR':                'Status',
    'Macbooks & iMacs':  'Status',
    'Active Servers':    'Status',
    'Decomm Servers':    'Status',
    'iPads':             'Status',
    'Kiosk Machines':    'Device Status',
    'iPads Macbooks to Decom': 'Device Status',
}
IT_STATUS_RED = {'broken','decommissioned','decomm','obsolete','dead','shutdown','migrated','broken screen/obsolete'}
IT_STATUS_AMBER = {'partially broken','partial','faulty','issue','to be decom','to decommission','decommission list'}

# For each category, which field's distinct values populate the left sidebar filter?
IT_SIDEBAR_FIELD = {
    'Active Servers':                   'Type of Server',
    'Decomm Servers':                   'Type of Server',
    'Network Devices':                  'Location',
    'IP Phones':                        'Brand',
    'CCTV':                             'Device Type',
    'CCTV Locations':                   'Location',
    'CCTV Models':                      'Type',
    'APs':                              'Manufacturer',
    'Printers':                         'Location',
    'UPS':                              'Model',
    'iPads':                            'Building',
    'VR':                               'Status',
    'Mobile Devices':                   'Department',
    'Chromebooks':                      'Category',
    'Macbooks & iMacs':                 'Location',
    'Desktops':                         'Building',
    'Laptops':                          'Category',
    'Projectors':                       'Building',
    'Cameras':                          'Building',
    'Infants Camera':                   'Brand',
    'Smartboards':                      'Building',
    'TV':                               'Class/Office',
    'Audio Visual devices':             'Location',
    'Access Control':                   'Location',
    'Kiosk Machines':                   'Device Status',
    'Monitors n Docking Station':       'School/Department',
    'PC Laptop to Decom':               'Type',
    'iPads Macbooks to Decom':          'Device Status',
    'Smartboard to Decom':              'Brand',
}


# ── THEMES ─────────────────────────────────────────────────
DARK = {
    'bg':       '#091829',
    'surface':  '#0F2540',
    'surface2': '#132944',
    'card':     '#0F2540',
    'sidebar':  '#091829',
    'topbar':   '#081730',
    'border':   '#1E3A5F',
    'gold':     '#C9972C',
    'gold2':    '#E8B84B',
    'text':     '#FFFFFF',
    'text2':    '#C2D0E6',
    'text3':    '#8892A4',
    'text4':    '#5A7A9A',
    'blue':     '#2563EB',
    'teal':     '#4ECCA3',
    'orange':   '#D4700A',
    'red':      '#E87060',
    'red_bg':   '#2A1A18',
    'in_bg':    '#0F2A1A', 'in_fg': '#4ECCA3',
    'low_bg':   '#2A1E08', 'low_fg': '#D4700A',
    'out_bg':   '#2A1210', 'out_fg': '#E87060',
    'row_alt':  '#0C2035',
    'side_txt': '#E8EEF7',
    'side_mut': '#6A8AAA',
    'check_bg': '#1A3A5C',
    'check_sel':'#C9972C',
}

LIGHT = {
    'bg':       '#EEF2F9',
    'surface':  '#FFFFFF',
    'surface2': '#F3F6FC',
    'card':     '#FFFFFF',
    'sidebar':  '#0B2A52',
    'topbar':   '#FFFFFF',
    'border':   '#E4EAF3',
    'gold':     '#C9972C',
    'gold2':    '#E8B84B',
    'text':     '#10213E',
    'text2':    '#3B4B6B',
    'text3':    '#6B7A96',
    'text4':    '#8493AE',
    'blue':     '#2563EB',
    'teal':     '#16A34A',
    'orange':   '#B45309',
    'red':      '#D02D2D',
    'red_bg':   '#FEF4F4',
    'in_bg':    '#E7F7EE', 'in_fg': '#137A3A',
    'low_bg':   '#FFF4E5', 'low_fg': '#B45309',
    'out_bg':   '#FDECEC', 'out_fg': '#D02D2D',
    'row_alt':  '#FAFBFE',
    'side_txt': '#FFFFFF',
    'side_mut': '#9DB0CC',
    'check_bg': '#E8F0FE',
    'check_sel':'#2563EB',
}

def T(): return DARK if SETTINGS.get('theme','dark')=='dark' else LIGHT

# ── DB CONFIG ──────────────────────────────────────────────
CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost;DATABASE=BSBStationery;"
    "UID=sa;PWD=sa123;TrustServerCertificate=yes;"
)
def get_conn(): return pyodbc.connect(CONN_STR)
def qy(sql, params=()):
    with get_conn() as c:
        cur = c.cursor(); cur.execute(sql, params)
        return [tuple(r) for r in cur.fetchall()]
def ex(sql, params=()):
    with get_conn() as c:
        cur = c.cursor(); cur.execute(sql, params); c.commit()

# ── EMAIL CONFIG ───────────────────────────────────────────
# To send REAL emails, fill these in and set 'enabled': True.
# For Gmail/Outlook use an "App Password" (not your normal password).
# Until enabled, the app shows the email content in a popup instead.
SMTP_CONFIG = {
    'enabled':     True,                       # <-- set True once filled in
    'host':        'smtp.gmail.com',            # e.g. smtp.office365.com
    'port':        587,
    'sender':      'mail.amrziyad@gmail.com',    # the "from" address
    'password':    'yekvoypydpbeimkp',                          # app password
    'system_link': 'http://localhost/bsb',      # link staff use to access the system
}
EMAIL_DOMAIN = 'thebsbh.com'

# Departments staff can belong to. These mirror the app's working areas.
DEPARTMENTS = ['Admin','IT','Procurement','Operations','Finance','HR',
               'Admissions','Marketing','Elements','Teaching','General']

def send_access_email(to_email, username, temp_password, parent=None, reset=False):
    """Send (or preview) the access-details email. Never logs the password to DB."""
    subject = 'BSB Resource Manager — Access Details'
    body = (
        f"Access Details:\n"
        f"Username: {username}\n"
        f"Temp Password: {temp_password}\n"
        f"Link For System: {SMTP_CONFIG.get('system_link','')}\n\n"
        f"{'Your password has been reset. ' if reset else 'Your account has been created. '}"
        f"Please log in and you will be asked to set your own password."
    )
    if SMTP_CONFIG.get('enabled'):
        try:
            import smtplib
            from email.mime.text import MIMEText
            msg = MIMEText(body)
            msg['Subject'] = subject
            msg['From'] = SMTP_CONFIG['sender']
            msg['To'] = to_email
            s = smtplib.SMTP(SMTP_CONFIG['host'], SMTP_CONFIG['port'], timeout=15)
            s.starttls(); s.login(SMTP_CONFIG['sender'], SMTP_CONFIG['password'])
            s.sendmail(SMTP_CONFIG['sender'], [to_email], msg.as_string()); s.quit()
            return True, f"Email sent to {to_email}"
        except Exception as e:
            return False, f"Could not send email: {e}"
    else:
        # Preview mode — return the content so the UI can show it
        return None, f"To: {to_email}\nSubject: {subject}\n\n{body}"

# ── LOGO LOADER (embedded, no external files needed) ───────
_LOGO_B64 = {
    'login':   "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",
    'sidebar': "iVBORw0KGgoAAAANSUhEUgAAADQAAAA6CAIAAABwEO3MAAASZklEQVR4nL2ZeWAURdbAq7p77ivnZHKSi5CEhCMBAgRCQIzIKSICoqDg6n6LAirKoou6KiuL53ogsOqihMtwCiHhyE1CBHKHJJBrksmcyUwmydzTXbV/tIwhXPH49v2TTHW9V79+3e/Vq9cQYwzuJOw4hPCOV/83KsRdFDCEEEKI0J3R76jFquBhqyD0swoA+I5aQ+FomoEQ7t+X+9jizTRNEwSkaea+yzAMAyHEGCGEIAGZ4akQBEQIrXj8ja/3nIR3WugWOJpmKIpsb+v6bm92cIj/grkvd6n0FEUyDEII3f78McYMw2CMSZJUqfQL522aP+clZbuGpEjPpduxEEIMg0iS1Ol6Fs172V/u/cPhvKZGJUWRQ/nwTXG7aYxxT3fv9MnPnjhWgDE+sD83JWnV0SN5njk0zdA0TbtpmmYQQp7x7/5zemLSqqNZeadOFk9MWvX17uOeSwghmmbc7p+1POOnThanJK/e++2PGOPcM6Wpk9ZoNT0eDFYAq8/eZU31jWkpa/d+cxJj3NTQ6na7W1tUC+ZsWPbYX8/mlNntDnyrdHUZvt5zYm7GhhXL3lCp9BgjxNBabc+qJ96c88CLe746purUDVFxOJwXzv30xONvzM1Y39jYjhDTUN+CMT6YmTN1wporlxswxjcfFIYYIQAhAODLz7P27c1+Z9vzGXOm6LQ92adKuFxq4uSE2NiI7FMlmd/l9PVZAoP8AhQ+FIfTbx7Qaoz9/daY2LCn1ywYnzSqtuZGTVUTACBhzMjxSXF1ta3/+ebHpoZ2kUgQFOwv85LQNG3QG7WaHrFEtOLJhx5ZnN7S3HmprMbhcD48b3pISEBB/tXXN3+5bHnGxldWsI8UIoTMvf1rVr8rkgh37npNKhW3tqgu/1Q3d/50oVCQc6bEbnPOSJ+gCPQzGc0N19rV6m63m/bylkRFBsfEjiBJsuFaa8XVJoXCZ9bsiRAS+XmXteqe8cmjEhJHIsTcuNHZ1qI2mQYoigwK8o0fHeHn79NtMBbkV3C51Nx501xu+tSJwuSJ8TGjwq1W+/p1H+q1pr2Zb/r6eVEAAw6HGhkT2tqmMeh7pVLxhfPlCYnRMpkEYbxw0UydtrviyrX+AVtAgG9wqHzkqFCKomxWR3e3Kfv0RavFFhTs/8jidIlUxL7Esx+cbLXYqiqbDu7PEQoFwaH+CYmRIpGAYZDZbK2rbdXpLovFgmnTxwUFyzHGXB43Oibs/NnymFHh3YZei8U2MiaUw+HcEhDHjxVMSlp9MDMXY3zoQK6yXY0xpmn6Zri4O5TaiquNF4urSoqqLpXWNlxrMxrNHnX2rR38D8bYZOxraGgrL6stKaoqKa68eqVBqdS4XG5PeCGEu1T6/fuyMcZHs/ImJa364dD5WwLCEyNKpSZt6p8yvz/jdrsPZOYgxCCEELplvdtlcAwOHhwUzUOFYd95hDBGhw7kOBzOI1l5qSlrW5o78aCAvS2V9PROS1nT09NbWlLZ2tKJMfakDM8/g/MI61qEBvuYZhjkceHAgA1jTLtpjLHT6fIsx1roUGoK8y/3mQemTV6r0XTjW1MJ5Ul4FEW63bSvr9f8hdMP7T+7es38ny7VR0aFMjSiOGTeuUs5OZfkct+nnp6r6tDuzzwboPBdtiJDKOS+9bfdW9/6U0He5eYbKoIkIiKCUqYkfvLR/g8/2Vhxpami4lpycnxomOKjHZmffv7Ke3//d1x8xJKls2kacThke5t67LhRR7LyZz+YEhjo53bTHM4vSLfsECRJYIzTZoyvqWkRi0U2uwMAACAAAIglYqNxoKb6hsnYJ5VJGAZbrY7L5fUKhdzca5FKRY8une3tKwsJlT+27EGpVGw226RSSc6ZMovFOT4pViwRqtU9zTdUba1asVh80yqwWGwyL3FVRVPajHEYY4K4heeWH+w2HBwit1jsDGJYfYKAAAAenysRi6JHhtVUN/v6yjo7dA31bfMXpiHESKQiHo8nFgsFfJ5AwBeLhRTF8fKSMAzatPnJDqX+0MHzfr5eHA7V3W3GEAiFPAAAJCAAgN3u+8yWkFA5hJBd685wrPAFPMwgxCAACYQQe5NdKp3FauHzKZIi21q7fP2k6bOSP/4gU6/rUasMJlMfTTMdSk2HUoMQ6jaY1Gp9f//AiaOFkyfHI4R0uh6HzZ46faxUzFe2qwe7AyFMM0goFNxO8ktAeF5Sg8G0/LEtLpfr9KliT5yaTObWlg51lw5jPNBv6exQY4zb27psVqtSqbZZ7RjjluZOdZceY+xyOtvbVA6Hs7e3v7qyCWNssVibb7QhRLe3danVes+Kp0+V0LR7xdLXWUV0a4TfwXO30QMAAAGJyoqmS6W1AAClUltcWP35pwclEuHRrHy91qjXG//5/rcRkUEUh/p697Fv/n2Sx+fxeFyt2lBf37LryywIoEwm+2D7dzKZKChIju9S4Q6R+8MhhAAAZ7JLc8+UjxkX097W9fbW3ekPTJiaOsZudzY1ddbVtfb1Wy6XNyGEX974cer0cfGjI19a/6leZ3xr657ZD6YQJPHm33ZJpeLiwhqSJBGDIBhWtXx/ODaEFz0ygy/g/3iyRKvpEQpFwcHy5ImjQ0IVPj4yZYfucnlDWJiCYRid1hwXH5mWnkzTTH1dK0mSAQq/jIemtLZoKA4ZEqqQyiQEef9FhwvHFtPnzpU98mjatfrWyOgQp9NRUlx1+sfia/UtBr1xbGLU3HlTmpraIYQxo0IOZOZknyqJHhk0ZWoiQqjsYnXumdLZGRNsVnt7W1dJceWF85cABHeoXX8DHHv4SEqK1+tMW99+VqHw+/yr1wx6E8Y4JDTg4XlTIqOCIYCbNq/kcjmffbHJX+7D43He+vtzQpHgi69eU6sNo2LD1724bGDAtuGVFX19lpAQBQDDeq7UfWewR6OwEYErnwoEAGCM5XKfJUsfYK/OmDkRANBr6uszDxQVViQmRveZzYmJMQIBHwDgL/dZuiyDnRkULA8Klg8x+3vhWGHjHEKCICC7mbqcTq3W6HK5vb0ln/3r8Oqn5x3LyrdaHNmnL/H5fF9/Wa9pgCAIRaCvQMBnUdizHEHAYR4fhwsHAAAAsrUCewKkOJRQyKcois/n+fjIGhuUIrGQokihgJc8IZ7D4QgEfIpDUhSFMYAQIIRJks3qw5XhwkEISfKX2yUIgiAIRaAf+/OVV58su1iVOm2Mn79PeHhgYJA/AEAmkwy2QJIQo6G7571lOEkYAwDqa5s//Oe+7/dm63Q9AIDjR/OsVrtKpdu+7Vtzb//+788AAOtqWz7Y8V2nSkfTNE0zWT+cBwBUXGn48rPDxYVXv/oyCxLwYOaZi8WV4OYj/r1wHqmpbi4qvFpVcZ1hmPff+7668rpc7pOfV+Fy0eU/1Xd393Z0GBw213++Od3ZoevpMb339297unt5fG55ef34pNhPPjp4cH+un7/3MLeHXwHn5S3h87lSqdhN09fqWvz8xSXFVTweVyqTFBdVDPTb5XLf8IjAkos1QQrfyKiQooIKPz9pUWGVItBPHuBHEPAf2/+vrLQ263B+aJjiD4bTaXscDsfohMi8c1cK8q5u276utKSmqrLRZrMvWfpAQIC3UqlpaVGNjo+02O0H9udoNca333nuh4PnGurb9DqjUqnpUunf3/FCdfUNi8U2zEWHGxB+/l5Ll82SSiVzHp5643pHXFzkhpeWu930ixuW6nXGefNTxRKB2+UeN3ZkcKi8pKhqWtq4uLjwZ9bO53LJtc/OJ0kyISFSLBYe//GfPB4H3MztvxeOzUnhESHhESHsSEhoAABg9kOTPXMUgf6DVR5f/nPinTNvmmcwNi4SYxwcohhs9vfC3U0QxohBGGO204MRBhBg/HOhQFEk+Ln7RLD1Ldu/ghBAONwe3nDhXC6XubdfIhWTBGGzORBCXl5iiqIAAAghFmWw2Gw2kiR5PB4A2Gq18XlcdvKvkmEoYIABdjpcBXkVu3aesFjtk1JiRQJBeXl9ZHTIptdWJiRG9/b2H9iXazT1yWSShYumUxQsK61paVZfOH9V5iUZOzaqtaWrr9/2yOK0Z9YuvNkwvL8MI1ohAABIpOJlTzwEIKitaU6flbzjkw0LHknbl5n74roPa2tuzMvYqFRqH56bGhomf3vr7gVzNyUnj06flVx2sV6p1G5+ffU/dqwz6Hv//NyOXTuPQAgZZlib2DBTCfvS0Hw+ly/gu10MTTNLHpsZGR6kUZtWPP4mQZEffLxh4qT4R5fM2rlnC5dDmc0DfB5PIhEKhTybzRkWFvjoknSBUJB9qgzcPNH9UXCAtcie48ViAUWRtTXNOq2JwyEZBpt6+utqmgEATqdbJBJs3rLKYrGxHsIY+/l5AQAaG5QDtoHRCRFgeHvXr4NjjfJ43KKCqm3vfLP+hU8y5qScOLV9RJjcaBpYtfLtzH05bHEwPX386MRoq9VBUaTD4d63N/v5tdvy8yu3vfvn7TteQAj98Z4DAEAIMMLePuKExKiUlLjGBuXxo8VTUhNcTpfF5npl42fz57x8NqdULvf19/d2Ol2QABBCHp8THh4klQoqrly/eLF6+IXJr4WDLrc7cezIRYvT93y7hSKJN17fLZYI/vyXRd0Gk0gkaGrqfGrlu6++/C+MEUkSDIN4PM7jyx/csvWZv7zw6Kkfy55bu71LpRvmR4S7wmGMhxT6bJ1IENBmdbhcbm9vr1GxYSKhoLVF/f6OF3ftfjUo0Mdud/r4SL/84ti/95wQi4XsGbm7u9flck+cNDo4xK+/39bcrLppf9DfYcBBAIBIxKNpxuWiCQIyzC+9d4riEATECIhEfC6Xo1Sq6+pabVb7mDFRCKHlK+eczfvsja1PI8TIvMQnj5XYbHYOlyIJwt/fm8vlVFxt0mmNfv6yuPgI9iYZhAAAbjdN04xQxL8d7pYkDCFAGAuFQopDajU9XjKhVtMTGqbAGDscrtKL1Z2dej6fm3U4T6M2HD50QSYVb97yZK+pnyAIp9MlEgvXb1wWHR307JrtZvPAmexSt5vu6jLs3nl0YMC2+6sTk6eOfm/b8wqFH8MggoAGvUkqEXYbeiEAXl5SjPGQ5AyHuJX9TvLxhwdcLtcrr67MOnz+iSfnAgAGBizXm5Q8LpfiUDarw+l0yrwkMTEjCIqclPT095lvjk6IZhgEISAIYnTMsoio4L9uWUUQEEBgGbABAMLCAmLjIwCADMOQJAkAOHQgd9Hi9D27jlst9te3rmFohhyyDQ5piLLNUIOhd9rkZw16U3VV4w+Hch0OJ76L9Jr6I8IWJ8QuO3+u3Ol00TSza+eR8OAFdbXNd1PBGLtcrmNZFy7/VGfu7U9NWaNWGxAa2sXBGA/1HACAYRBJEkd+uLDziyPZZz/VqA0XS6qDguRR0aHBIXK2GgM3q4yebvOhg7mJY0bmnikDGCGErVbH+peWx8VH0m4aAAgg8JQFNM1o1IbmG6quLt3kKYmRUSHzH9q46pl5K5+aixh0e5viDnAevt07j+7bl/PRxxtTpiR0dmqbGpXGHrNAwAsJCYiICvb19RqiZbfbMcJCkdBjgR03mfra27pUKr3D5vL2lcTGho8ID6682vjSxk+WLJm5/qUVgyffH85jvbDg6vvbvpP7e69YmTFj5gSRiO9wOK83Ktvb1A6nK0DhGxsbzh4EEUJsdvWspNcZmxrbtVojj0uFRwbFxkUIBHy73VFUWHlw/3mNxvDaX596MGPy3cjuBTd4mePHCk+fLFF36QMUvqnTx856YEL0yDAAgFKpqa9t7u+3xowakZQcdzP146rK602N7SKRMGFMVGRkCACgtUVVkF9RWlKt0xoVgX7zFkx7dOksAsJ7kN0HbrA/AABmc/9Pl64VF1Vdq2+jaXrUqBFpM8bPnD1BKhWVFFW1tnY9MHsSl8c5e6Y0JCwwfWaSzeooyK8oLqi43tQJKRgfHzk9bdyUqYnePrLbjf8WOI8LAQCDb7GlubP8Un3Zxdqm652pqWPe2fY87ab3fX/K6aJXr14oFPHfffvrwoKKmJiwqamJk6cmjowZMcQaQRD3rzjvEfC3Zxn2e+vgwZ5u899e/yo1ZU1jQ5vL5bZabe3t6rSpf9q86XO93jh4JtsGuD1f3EN+Bdw9QM+dLZ8w5sn2ti6DwTRh3FMnjxfeBGJ+LdAfADeYkv3QVpB3JWPmuoxZ606eKMIYu1zu38zkkd9+NGQFQsjhUG43nT5rQmlpba+pf+GitCFfiX67cTzstso9BGOAMaJpGiHM5XKHfzL9X8D9P8l/AU8RE/tsrH7aAAAAAElFTkSuQmCC",
    'topbar':  "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",
}
_logos = {}
def load_logo(name, path, size):
    try:
        import base64, io
        from PIL import Image as PILImage, ImageTk
        b64 = _LOGO_B64.get(name)
        if not b64: return None
        pil = PILImage.open(io.BytesIO(base64.b64decode(b64))).convert('RGB').resize(size, PILImage.LANCZOS)
        photo = ImageTk.PhotoImage(pil)
        _logos[name] = photo
        return photo
    except: return None

# ── FONTS ──────────────────────────────────────────────────
F = {}
def init_fonts():
    CG = 'Cormorant Garamond'  # login/headers
    OT = 'Outfit'              # body (web app font)
    # Fall back gracefully if not installed
    F.update({
        'h1':     ctk.CTkFont(CG, 24, 'bold'),
        'h2':     ctk.CTkFont(CG, 19, 'bold'),
        'h3':     ctk.CTkFont(OT, 15, 'bold'),
        'body':   ctk.CTkFont(OT, 14),
        'body_b': ctk.CTkFont(OT, 14, 'bold'),
        'sm':     ctk.CTkFont(OT, 13),
        'sm_b':   ctk.CTkFont(OT, 13, 'bold'),
        'xs':     ctk.CTkFont(OT, 11, 'bold'),
        'tiny':   ctk.CTkFont(OT, 10),
        'login':  ctk.CTkFont(CG, 26, 'bold'),
        'disp':   ctk.CTkFont(CG, 32, 'bold'),
        'caps':   ctk.CTkFont(OT, 10, 'bold'),
        'mono':   ctk.CTkFont('Consolas', 13),
    })

# ── TREE STYLE (themed) ────────────────────────────────────
def init_tree():
    C = T()
    s = ttk.Style(); s.theme_use('clam')
    s.configure('BSB.Treeview',
                background=C['card'], foreground=C['text'],
                fieldbackground=C['card'], rowheight=40,
                font=('Outfit', 13, 'bold'), borderwidth=0)
    s.configure('BSB.Treeview.Heading',
                background=C['surface2'], foreground=C['text4'],
                font=('Outfit', 11, 'bold'), relief='flat', padding=10)
    s.map('BSB.Treeview',
          background=[('selected', C['blue'] if SETTINGS.get('theme')=='light' else '#1A3A5C')],
          foreground=[('selected', C['text'])])

def mktree(parent, cols, widths, height=15):
    C = T()
    h = tk.Frame(parent, bg=C['card'])
    t = ttk.Treeview(h, columns=cols, show='headings',
                     selectmode='extended', height=height, style='BSB.Treeview')
    for c, w in zip(cols, widths):
        t.heading(c, text=c)
        t.column(c, width=w, minwidth=40, anchor='w')
    t.tag_configure('low', background=C['low_bg'], foreground=C['low_fg'])
    t.tag_configure('out', background=C['out_bg'], foreground=C['out_fg'])
    t.tag_configure('alt', background=C['row_alt'])
    t.tag_configure('in',  background=C['card'],   foreground=C['text'])
    vsb = ttk.Scrollbar(h, orient='vertical', command=t.yview)
    t.configure(yscrollcommand=vsb.set)
    t.pack(side='left', fill='both', expand=True)
    vsb.pack(side='left', fill='y')
    return h, t

def status_info(qty, mn):
    if qty <= 0:  return 'out', '● Out'
    if qty < mn:  return 'low', '▲ Low'
    return 'in', '✓ OK'

# ── WIDGET HELPERS ─────────────────────────────────────────
def L(p, text, color=None, fk='body', **kw):
    C = T()
    return ctk.CTkLabel(p, text=text, text_color=color or C['text'], font=F[fk], **kw)

def caps_lbl(p, text):
    C = T()
    return ctk.CTkLabel(p, text=text.upper(), text_color=C['text4'], font=F['caps'])

def blue_btn(p, text, cmd, w=140, h=38):
    C = T()
    return ctk.CTkButton(p, text=text, command=cmd,
                         fg_color=C['blue'], hover_color='#1746B0',
                         text_color='white', font=F['sm_b'],
                         corner_radius=6, width=w, height=h)

def gold_btn(p, text, cmd, w=140, h=38):
    C = T()
    return ctk.CTkButton(p, text=text, command=cmd,
                         fg_color=C['gold'], hover_color=C['gold2'],
                         text_color='#0B1F3A', font=F['sm_b'],
                         corner_radius=6, width=w, height=h)

def ghost_btn(p, text, cmd, color=None, w=120, h=34):
    C = T()
    fc = color or C['text3']
    return ctk.CTkButton(p, text=text, command=cmd,
                         fg_color=C['surface2'], hover_color=C['border'],
                         text_color=fc, font=F['sm_b'],
                         corner_radius=6, width=w, height=h,
                         border_width=1, border_color=C['border'])

def red_btn(p, text, cmd, w=110, h=34):
    C = T()
    return ctk.CTkButton(p, text=text, command=cmd,
                         fg_color=C['red_bg'], hover_color=C['out_bg'],
                         text_color=C['red'], font=F['sm_b'],
                         corner_radius=6, width=w, height=h,
                         border_width=1, border_color=C['out_fg'])

def inp(p, placeholder='', w=320, show=None):
    C = T()
    kw = dict(placeholder_text=placeholder, width=w, height=40,
              fg_color=C['surface2'], border_color=C['border'],
              text_color=C['text'], placeholder_text_color=C['text4'],
              font=F['body'], corner_radius=6, border_width=1)
    if show: kw['show'] = show
    return ctk.CTkEntry(p, **kw)

def hline(p, pady=4):
    C = T()
    ctk.CTkFrame(p, height=1, fg_color=C['border']).pack(fill='x', pady=pady)

# ── CHECKBOX TREE (Gmail-style using ttk.Treeview) ────────
class CheckboxTree:
    """Treeview with checkbox column - reliable ttk-based implementation"""
    def __init__(self, parent, cols, widths, height=15, order_mode=False, image_col=False, simple_mode=False):
        C = T()
        self.checked = set()
        self.all_ids = []
        self.order_mode = order_mode
        self.simple_mode = simple_mode       # generic checkbox list (e.g. categories) — no stock logic
        self.image_col = image_col           # show a per-row 🖼 button column
        self.on_image_click = None           # callback(item_id) when 🖼 clicked
        self._id_map = {}  # tree_item -> item_id
        self._val_map = {} # item_id -> values
        
        self.frame = tk.Frame(parent, bg=C['card'])
        
        # Build columns: checkbox first, then optional image column
        extra = ('_img',) if image_col else ()
        all_cols = ('_cb',) + extra + tuple(cols)
        
        tree_frame = tk.Frame(self.frame, bg=C['card'])
        tree_frame.pack(fill='both', expand=True)
        
        self.tree = ttk.Treeview(tree_frame, columns=all_cols, show='headings',
                                  selectmode='extended', height=height, style='BSB.Treeview')
        
        # Checkbox column header
        self.tree.heading('_cb', text='☐', command=self._toggle_all)
        self.tree.column('_cb', width=36, minwidth=36, stretch=False, anchor='center')

        if image_col:
            self.tree.heading('_img', text='Image')
            self.tree.column('_img', width=84, minwidth=84, stretch=False, anchor='center')
        
        for c_name, w in zip(cols, widths):
            self.tree.heading(c_name, text=c_name)
            self.tree.column(c_name, width=w, minwidth=30)
        
        # Tags for status
        C2 = T()
        self.tree.tag_configure('low', background=C2['low_bg'], foreground=C2['low_fg'])
        self.tree.tag_configure('out', background=C2['out_bg'], foreground=C2['out_fg'])
        self.tree.tag_configure('alt', background=C2['row_alt'])
        self.tree.tag_configure('in',  background=C2['card'],   foreground=C2['text'])
        self.tree.tag_configure('checked_low', background=C2['low_bg'], foreground=C2['low_fg'])
        self.tree.tag_configure('checked_out', background=C2['out_bg'], foreground=C2['out_fg'])
        self.tree.tag_configure('checked', background=C2['surface2'], foreground=C2['text'])
        # Status row colors (broken / partially broken / decommissioned)
        self.tree.tag_configure('status_red', background=C2['out_bg'], foreground=C2['out_fg'])
        self.tree.tag_configure('status_amber', background=C2['low_bg'], foreground=C2['low_fg'])
        
        vsb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview,
                            style='BSB.Vertical.TScrollbar')
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='left', fill='y')
        
        # Click to toggle checkbox
        self.tree.bind('<Button-1>', self._on_click)
        self.tree.bind('<Double-1>', self._on_dbl)
        self._dbl_callback = None

    def _on_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        col = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if not item:
            return
        if col == '#1':  # checkbox column
            self._toggle_item(item)
        elif self.image_col and col == '#2':  # image button column
            iid = self._id_map.get(item)
            if iid is not None and self.on_image_click:
                self.on_image_click(iid)
        
    def _on_dbl(self, event):
        if self._dbl_callback:
            self._dbl_callback(event)

    def _toggle_item(self, tree_item):
        iid = self._id_map.get(tree_item)
        if iid is None: return
        if iid in self.checked:
            self.checked.discard(iid)
            self._set_check(tree_item, False)
        else:
            self.checked.add(iid)
            self._set_check(tree_item, True)
        self._update_hdr()

    def _set_check(self, tree_item, checked):
        vals = list(self.tree.item(tree_item)['values'])
        vals[0] = '☑' if checked else '☐'
        self.tree.item(tree_item, values=vals)

    def _toggle_all(self):
        if len(self.checked) == len(self.all_ids):
            # uncheck all
            self.checked.clear()
            for ti, iid in [(ti, self._id_map[ti]) for ti in self.tree.get_children()]:
                self._set_check(ti, False)
            self.tree.heading('_cb', text='☐')
        else:
            # check all
            self.checked = set(self.all_ids)
            for ti in self.tree.get_children():
                self._set_check(ti, True)
            self.tree.heading('_cb', text='☑')

    def _update_hdr(self):
        n = len(self.checked); total = len(self.all_ids)
        if n == 0: self.tree.heading('_cb', text='☐')
        elif n == total: self.tree.heading('_cb', text='☑')
        else: self.tree.heading('_cb', text='☒')

    def clear(self):
        self.tree.delete(*self.tree.get_children())
        self._id_map.clear(); self._val_map.clear()
        self.checked.clear(); self.all_ids.clear()
        self.tree.heading('_cb', text='☐')

    def insert(self, item_id, values, tag=''):
        if self.simple_mode:
            # Generic checkbox list — values are shown as-is, no stock logic.
            # Used for things like category lists (Name, Parent, item count...).
            display = ('☐',) + tuple(values)
            tag_name = tag or ('alt' if len(self.all_ids) % 2 == 0 else 'in')
            ti = self.tree.insert('', 'end', values=display, tags=(tag_name,))
            self._id_map[ti] = item_id
            self._val_map[item_id] = values
            self.all_ids.append(item_id)
            return
        if self.order_mode:
            # values = (iid, name, cat, unit) — NO stock shown to staff
            iid, name, cat, unit = values
            if self.image_col:
                display = ('☐', '🖼  View', iid, name, cat, unit)
            else:
                display = ('☐', iid, name, cat, unit)
            tag_name = 'alt' if len(self.all_ids) % 2 == 0 else 'in'
            ti = self.tree.insert('', 'end', values=display, tags=(tag_name,))
            self._id_map[ti] = item_id
            self._val_map[item_id] = values
            self.all_ids.append(item_id)
            return
        # values may be 7-tuple (no LPO/Invoice) or 9-tuple (+LPO,+Invoice)
        if len(values) >= 9:
            iid, name, cat, qty, unit, price, mn, lpo, inv = values[:9]
        else:
            iid, name, cat, qty, unit, price, mn = values
            lpo, inv = '', ''
        st, slabel = status_info(qty, mn)
        stock_str = f"{qty}  {slabel}"
        if self.image_col:
            display = ('☐', '🖼  View', iid, name, cat, stock_str, unit, f'{float(price):.3f}', mn, lpo or '—', inv or '—')
        else:
            display = ('☐', iid, name, cat, stock_str, unit, f'{float(price):.3f}', mn, lpo or '—', inv or '—')
        
        tag_name = st if st != 'in' else ('alt' if len(self.all_ids) % 2 == 0 else 'in')
        ti = self.tree.insert('', 'end', values=display, tags=(tag_name,))
        self._id_map[ti] = item_id
        self._val_map[item_id] = values
        self.all_ids.append(item_id)

    def get_selected_values(self):
        return [self._val_map[i] for i in self.checked if i in self._val_map]

# ════════════════════════════════════════════════════════════
class BSBApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        init_fonts(); init_tree()
        self._apply_theme()
        self.title('BSB Resource Manager')
        self.state('zoomed')  # maximize on startup
        self.resizable(True, True)
        self.current_user = None
        self.current_role = None
        self.current_username = None
        self.current_email = None
        self._notif_seen_orders = set()
        self._notif_seen_msgs = set()
        self._notif_started = False
        self.cart = []
        self._splash()

    # ══ CUSTOM OPENING ANIMATION ═════════════════════════════
    def _splash(self):
        """Polished animated intro → dissolves seamlessly into login (no white flash)."""
        NAV   = '#06122B'   # deep brand navy
        NAV_L = '#0B2A52'
        GOLD  = '#C9972C'
        GOLD2 = '#E8B84B'
        WHITE = '#FFFFFF'
        MUT   = '#6A89AD'
        BARBG = '#15294A'

        self.title('BSB Resource Manager')
        try: self.state('zoomed')
        except: pass
        self.configure(fg_color=NAV)

        splash = tk.Frame(self, bg=NAV)
        splash.place(relx=0, rely=0, relwidth=1, relheight=1)

        # Background canvas for subtle floating particles (depth)
        canvas = tk.Canvas(splash, bg=NAV, highlightthickness=0, bd=0)
        canvas.place(relx=0, rely=0, relwidth=1, relheight=1)

        # Thin gold accent strips top & bottom
        tk.Frame(splash, bg=GOLD, height=3).place(relx=0, rely=0, relwidth=1)
        tk.Frame(splash, bg=GOLD, height=3).place(relx=0, rely=1.0, anchor='sw', relwidth=1)

        center = tk.Frame(splash, bg=NAV)
        center.place(relx=0.5, rely=0.5, anchor='center')

        self._splash_jobs = []
        def after(ms, fn):
            jid = self.after(ms, fn); self._splash_jobs.append(jid); return jid

        def lerp(c1, c2, t):
            a = tuple(int(c1[i:i+2],16) for i in (1,3,5))
            b = tuple(int(c2[i:i+2],16) for i in (1,3,5))
            r = tuple(int(a[i]+(b[i]-a[i])*max(0,min(1,t))) for i in range(3))
            return f'#{r[0]:02x}{r[1]:02x}{r[2]:02x}'

        # ── Floating particle field (subtle, premium) ──
        import random
        self._splash_alive = True
        particles = []
        def seed_particles():
            try:
                w = self.winfo_width() or 1400
                h = self.winfo_height() or 850
            except: w, h = 1400, 850
            for _ in range(40):
                x = random.randint(0, w); y = random.randint(0, h)
                r = random.choice([1, 1, 2, 2, 3])
                spd = random.uniform(0.15, 0.6)
                shade = random.choice([NAV_L, '#1B3A6A', GOLD if r==1 else NAV_L])
                dot = canvas.create_oval(x-r, y-r, x+r, y+r, fill=shade, outline='')
                particles.append([dot, spd, h])
        def drift():
            if not self._splash_alive: return
            for p in particles:
                dot, spd, h = p
                canvas.move(dot, 0, -spd)
                coords = canvas.coords(dot)
                if coords and coords[3] < 0:
                    nx = random.randint(0, self.winfo_width() or 1400)
                    canvas.coords(dot, nx, h, nx+2, h+2)
            after(40, drift)

        # ── Logo with scale-in pop ──
        self._splash_imgs = {}
        def make_logo(size):
            img = load_logo(f'sp{size}', 'bsb_logo_login.png', (size, int(size*1.12)))
            if not img:
                img = load_logo('login', 'bsb_logo_login.png', (size, int(size*1.12)))
            self._splash_imgs[size] = img
            return img

        logo_lbl = tk.Label(center, bg=NAV, bd=0, highlightthickness=0)
        base_img = make_logo(150)
        if not base_img:
            logo_lbl = tk.Label(center, text='BSB', bg=GOLD, fg=NAV,
                                font=('Cormorant Garamond', 40, 'bold'), padx=24, pady=10)
        logo_lbl.pack(pady=(0, 6))

        def logo_pop(sizes=(70, 96, 120, 138, 150, 156, 150), i=0):
            if not base_img:
                return
            if i >= len(sizes):
                logo_lbl.configure(image=self._splash_imgs.get(150, base_img)); return
            s = sizes[i]
            img = self._splash_imgs.get(s) or make_logo(s)
            try: logo_lbl.configure(image=img)
            except: pass
            after(34, lambda: logo_pop(sizes, i+1))

        # ── Text stack (revealed on timeline) ──
        bsb_lbl = tk.Label(center, text='BSB', bg=NAV, fg=NAV,
                           font=('Cormorant Garamond', 44, 'bold'))
        # expanding gold divider under BSB
        divider = tk.Frame(center, bg=GOLD, height=2)
        name_lbl = tk.Label(center, text='Resource Manager', bg=NAV, fg=NAV,
                            font=('Cormorant Garamond', 27, 'bold'))
        sub_lbl = tk.Label(center, text='B R I T I S H   S C H O O L   O F   B A H R A I N',
                          bg=NAV, fg=NAV, font=('Outfit', 10, 'bold'))

        load_wrap = tk.Frame(center, bg=NAV)
        bar_row = tk.Frame(load_wrap, bg=NAV)
        bar_bg = tk.Frame(bar_row, bg=BARBG, height=4, width=300)
        bar_fg = tk.Frame(bar_bg, bg=GOLD, height=4, width=0)
        pct_lbl = tk.Label(bar_row, text='0%', bg=NAV, fg=MUT, font=('Outfit', 10, 'bold'), width=5)
        status_lbl = tk.Label(load_wrap, text='Initialising', bg=NAV, fg=MUT, font=('Outfit', 10))

        def fade_in(widget, target_fg, steps=12, step=0, delay=20):
            if not self._splash_alive: return
            if step > steps:
                widget.configure(fg=target_fg); return
            widget.configure(fg=lerp(NAV, target_fg, step/steps))
            after(delay, lambda: fade_in(widget, target_fg, steps, step+1, delay))

        def slide_in(widget, packkw, target_fg):
            widget.pack(**packkw)
            fade_in(widget, target_fg)

        # ── Timeline ──
        def show_bsb():
            slide_in(bsb_lbl, dict(pady=(16,0)), WHITE)
            after(180, grow_divider)

        def grow_divider(w=0):
            if not self._splash_alive: return
            if w == 0: divider.pack(pady=(6,0))
            if w >= 150:
                divider.configure(width=150); return
            divider.configure(width=w)
            after(8, lambda: grow_divider(w+10))

        def show_name():
            slide_in(name_lbl, dict(pady=(8,0)), GOLD)

        def show_sub():
            slide_in(sub_lbl, dict(pady=(8,0)), MUT)

        STATUSES = ['Initialising', 'Loading modules', 'Connecting database',
                    'Preparing inventory', 'Almost ready']
        def show_loading():
            load_wrap.pack(pady=(34,0))
            bar_row.pack()
            bar_bg.pack(side='left')
            bar_fg.place(x=0, y=0)
            pct_lbl.pack(side='left', padx=(12,0))
            status_lbl.pack(pady=(10,0))
            fade_in(pct_lbl, MUT); fade_in(status_lbl, MUT)
            animate_bar()

        def animate_bar(w=0):
            if not self._splash_alive: return
            if w >= 300:
                bar_fg.configure(width=300); pct_lbl.configure(text='100%')
                status_lbl.configure(text='Ready')
                after(320, finish)
                return
            bar_fg.configure(width=w)
            pct = int(w/300*100)
            pct_lbl.configure(text=f'{pct}%')
            idx = min(len(STATUSES)-1, pct//22)
            status_lbl.configure(text=STATUSES[idx])
            # ease-out: slow down near the end
            inc = 6 if pct < 80 else 3
            after(12, lambda: animate_bar(w+inc))

        # ── Seamless transition into login (NO white flash) ──
        def finish():
            # Briefly fade the splash content to solid navy (so the logo/text
            # gently vanish), then build the login behind a navy cover and lift
            # it away. The eye only ever sees navy → login, never white.
            fade_content_out()

        def fade_content_out(step=0, steps=8):
            if not self._splash_alive: return
            if step > steps:
                swap_to_login()
                return
            t = step/steps
            # fade every text element back toward navy (disappear)
            for wdg, col in [(bsb_lbl, WHITE), (name_lbl, GOLD), (sub_lbl, MUT),
                             (pct_lbl, MUT), (status_lbl, MUT)]:
                try: wdg.configure(fg=lerp(col, NAV, t))
                except: pass
            try:
                bar_fg.configure(bg=lerp(GOLD, NAV, t))
                bar_bg.configure(bg=lerp(BARBG, NAV, t))
                divider.configure(bg=lerp(GOLD, NAV, t))
            except: pass
            self.after(26, lambda: fade_content_out(step+1, steps))

        def swap_to_login():
            self._splash_alive = False
            for jid in self._splash_jobs:
                try: self.after_cancel(jid)
                except: pass
            # Build login behind a full navy cover, then drop the cover instantly.
            self._apply_theme()
            try:
                cover = tk.Frame(self, bg=NAV)
                cover.place(relx=0, rely=0, relwidth=1, relheight=1)
            except: cover = None
            try: splash.destroy()
            except: pass
            self._login()
            if cover is not None:
                try: cover.lift()
                except: pass
                def drop_cover():
                    try: cover.lower()
                    except: pass
                    try: cover.destroy()
                    except: pass
                self.after(140, drop_cover)

        # start particles after the window has a real size
        after(60, lambda: (seed_particles(), drift()))
        after(120,  logo_pop)
        after(620,  show_bsb)
        after(1150, show_name)
        after(1600, show_sub)
        after(2150, show_loading)

        def skip(_=None):
            self._splash_alive = False
            for jid in self._splash_jobs:
                try: self.after_cancel(jid)
                except: pass
            try: splash.destroy()
            except: pass
            self._apply_theme(); self._login()
        splash.bind('<Button-1>', skip)
        canvas.bind('<Button-1>', skip)
        self.bind('<Escape>', skip)

    def _apply_theme(self):
        C = T()
        if SETTINGS.get('theme','dark') == 'dark':
            ctk.set_appearance_mode('dark')
        else:
            ctk.set_appearance_mode('light')
        self.configure(fg_color=C['bg'])

    def _clear(self):
        for w in self.winfo_children(): w.destroy()

    def _reload(self, go_to=None):
        """Reload with new theme"""
        init_tree()
        self._apply_theme()
        self._clear()
        if go_to: go_to()
        elif self.current_user: self._modes()
        else: self._login()

    # ══ LOGIN (never changes with theme) ═════════════════════
    def _login(self):
        self._clear(); self.cart = []
        # Login is always the light split-panel design (never changes with theme)
        NAV  = '#0B2A52'  # left navy panel
        CARD = '#FFFFFF'  # right white card
        BG   = '#EEF2F9'  # page bg
        BLUE = '#2563EB'
        GOLD = '#C9972C'
        TXT  = '#10213E'
        TXT3 = '#6B7A96'
        TXT4 = '#8493AE'
        BD   = '#E4EAF3'
        SF2  = '#F3F6FC'
        RED  = '#D02D2D'
        SMUT = '#9DB0CC'
        STXT = '#FFFFFF'

        outer = ctk.CTkFrame(self, fg_color=BG)
        outer.place(relx=0, rely=0, relwidth=1, relheight=1)

        # Left navy panel
        left = ctk.CTkFrame(outer, fg_color=NAV, corner_radius=0, width=420)
        left.pack(side='left', fill='y'); left.pack_propagate(False)
        lc = ctk.CTkFrame(left, fg_color='transparent')
        lc.place(relx=0.5, rely=0.5, anchor='center')

        _img = load_logo('login', 'bsb_logo_login.png', (110, 124))
        if _img:
            tk.Label(lc, image=_img, bg=NAV, bd=0, highlightthickness=0).pack(pady=(0,16))
        else:
            logo = ctk.CTkFrame(lc, width=72, height=72, corner_radius=18, fg_color=BLUE)
            logo.pack(pady=(0,20)); logo.pack_propagate(False)
            ctk.CTkLabel(logo, text='BSB', text_color=CARD,
                        font=ctk.CTkFont('Segoe UI', 22, 'bold')).place(relx=0.5, rely=0.5, anchor='center')

        ctk.CTkLabel(lc, text='BSB Resource Manager', text_color=STXT,
                    font=ctk.CTkFont('Segoe UI', 22, 'bold')).pack()
        ctk.CTkLabel(lc, text='British School of Bahrain', text_color=SMUT,
                    font=ctk.CTkFont('Segoe UI', 13)).pack(pady=(4,0))
        ctk.CTkFrame(lc, height=1, fg_color='#1B3A6A', width=260).pack(pady=28)
        for txt in ['✓  Manage Inventory',
                    '✓  Order Stationery',
                    '✓  Assest Management',
                    '✓  Generate Excel reports']:
            ctk.CTkLabel(lc, text=txt, text_color=SMUT,
                        font=ctk.CTkFont('Segoe UI', 13)).pack(anchor='w', pady=4)

        # Right white login card
        right = ctk.CTkFrame(outer, fg_color=BG); right.pack(side='left', fill='both', expand=True)
        card = ctk.CTkFrame(right, fg_color=CARD, corner_radius=18,
                           border_width=1, border_color=BD)
        card.place(relx=0.5, rely=0.5, anchor='center')
        inner = ctk.CTkFrame(card, fg_color=CARD); inner.pack(padx=48, pady=44)

        ctk.CTkLabel(inner, text='Welcome back', text_color=TXT,
                    font=ctk.CTkFont('Segoe UI', 28, 'bold')).pack(anchor='w')
        ctk.CTkLabel(inner, text='Sign in to continue', text_color=TXT3,
                    font=ctk.CTkFont('Segoe UI', 13)).pack(anchor='w', pady=(4,28))

        ctk.CTkLabel(inner, text='USERNAME', text_color=TXT4,
                    font=ctk.CTkFont('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0,5))
        self._u = ctk.CTkEntry(inner, placeholder_text='Enter your username', width=360,
                              height=40, fg_color=SF2, border_color=BD, text_color=TXT,
                              placeholder_text_color=TXT4, font=ctk.CTkFont('Segoe UI',14),
                              corner_radius=10, border_width=1)
        self._u.pack(pady=(0,16))

        ctk.CTkLabel(inner, text='PASSWORD', text_color=TXT4,
                    font=ctk.CTkFont('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0,5))
        self._p = ctk.CTkEntry(inner, placeholder_text='Enter your password', width=360,
                              height=40, fg_color=SF2, border_color=BD, text_color=TXT,
                              placeholder_text_color=TXT4, font=ctk.CTkFont('Segoe UI',14),
                              corner_radius=10, border_width=1, show='●')
        self._p.pack(pady=(0,8))

        self._err = ctk.CTkLabel(inner, text='', text_color=RED, font=ctk.CTkFont('Segoe UI',12))
        self._err.pack(pady=(4,14))

        ctk.CTkButton(inner, text='Sign In', command=self._do_login,
                     fg_color=BLUE, hover_color='#1746B0', text_color='white',
                     font=ctk.CTkFont('Segoe UI', 14, 'bold'),
                     corner_radius=10, width=360, height=44).pack()



        self._u.focus()
        self.bind('<Return>', lambda e: self._do_login())


    def _do_login(self):
        u = self._u.get().strip(); p = self._p.get().strip()
        rows = qy("SELECT Role, FullName, Username, ISNULL(Email,''), ISNULL(MustChangePassword,0) FROM Users WHERE (Username=? OR Email=?) AND PasswordHash=?", (u,u,p))
        if rows:
            self.current_role = rows[0][0]
            self.current_user = rows[0][1] or rows[0][2]
            self.current_username = rows[0][2]
            self.current_email = rows[0][3]
            try: ex("UPDATE Users SET LastLogin=GETDATE() WHERE Username=?",(rows[0][2],))
            except: pass
            if rows[0][4]:
                self._force_change_pw()
            else:
                self._modes()
        else:
            self._err.configure(text='  ✕  Invalid username or password')
            self._err.pack(fill='x', pady=(6,0))

    # ══ MODE SELECT ══════════════════════════════════════════
    def _modes(self):
        self.state('zoomed')
        C = T()
        self._clear()
        self.configure(fg_color=C['bg'])
        
        # Settings bar at top
        self._make_settings_bar(self, lambda: self._reload(self._modes))

        role = self.current_role

        # Build mode list first so we know how many cards there are
        if role == 'admin':
            modes = [
                ('🏢','Procurement','Inventory & issue functions', self._dept_procurement),
                ('💻','IT','Manage & assign devices', self._dept_it),
                ('⚙️','Operation','Assign & manage functions', self._dept_operation),
                ('👥','Staff','General staff management', self._dept_staff),
                ('🧑‍💼','HR','HR management & reports', self._dept_hr),
                ('💰','Finance','Finance & payment management', self._dept_fin),
                ('🛡','Admin Dept','Staff history & oversight', self._dept_admin),
                ('👥','Users','Add, remove &\nmanage accounts', self._user_admin),
            ]
        elif role == 'procurement':
            modes = [
                ('📦','Manage','Add, edit & set quantity', self._manage),
                ('🛒','Orders','See & arrange staff orders', self._orders_mail),
                ('📊','Report','Generate & export reports', self._report),
            ]
        elif role == 'it':
            modes = [
                ('🔧','Manage','Manage IT assets & inventory', self._it_manage),
                ('📤','Assign','Assign devices to users', self._it_assign),
                ('🗑','Clearance','Equipment clearance & disposal', self._under_construction),
            ]
        elif role == 'operation':
            modes = [
                ('📋','Assign','Assign operational tasks', self._under_construction),
                ('⚙️','Manage','Manage operational assets (PPE register)', self._op_manage),
            ]
        elif role == 'hr':
            modes = [
                ('📦','Manage','HR records management', self._under_construction),
                ('📊','Report','Generate & export reports', self._under_construction),
            ]
        elif role == 'fin':
            modes = [
                ('💳','Payment Plan','Manage payment plans', self._fin_payment_plan),
            ]
        else:
            modes = [
                ('🛒','Order','Order stationery for your class', self._order),
                ('🕘','Order History','See your order history & status', self._order_history),
                ('🔑','Password','Change your password', self._change_pw_mode),
            ]

        # Scrollable outer frame so content never runs off-screen, regardless of
        # how many cards there are (admin has 7 → 3 rows + title + sign out).
        outer = ctk.CTkScrollableFrame(self, fg_color=C['bg'])
        outer.pack(fill='both', expand=True)
        center = ctk.CTkFrame(outer, fg_color=C['bg'])
        center.pack(pady=(36, 24))

        L(center, f'Good day, {self.current_user}', C['text'], 'disp').pack(pady=(0,4))
        L(center, self.current_role.upper() + ' ACCOUNT', C['gold'], 'caps').pack(pady=(0,36))

        grid = ctk.CTkFrame(center, fg_color=C['bg']); grid.pack()
        cur_row = None
        for idx, (icon, title, desc, cmd) in enumerate(modes):
            if idx % 3 == 0:
                cur_row = ctk.CTkFrame(grid, fg_color=C['bg']); cur_row.pack(pady=8)
            card = ctk.CTkFrame(cur_row, fg_color=C['surface'], corner_radius=10,
                               border_width=1, border_color=C['border'], width=200, height=180)
            card.pack(side='left', padx=12); card.pack_propagate(False)
            L(card, icon, fk='disp').pack(pady=(20,6))
            L(card, title, C['text'], 'h3').pack()
            L(card, desc, C['text3'], 'sm').pack(pady=(6,0), padx=14)
            if title == 'Orders':
                pend = (qy("SELECT COUNT(*) FROM Orders WHERE Status='Pending'") or [(0,)])[0][0]
                if pend:
                    ctk.CTkLabel(card, text=f' {pend} new ', text_color='white',
                                 fg_color=C['red'], font=F['caps'], corner_radius=8).pack(pady=(6,0))
            for w in [card]+list(card.winfo_children()):
                w.bind('<Button-1>', lambda e, c2=cmd: c2())
                w.bind('<Enter>', lambda e, w=card: w.configure(border_color=C['gold']))
                w.bind('<Leave>', lambda e, w=card: w.configure(border_color=C['border']))

        hline(center, 24)
        ghost_btn(center, '⬅  Sign Out', self._login, w=170).pack(pady=(0, 20))

        if role in PROCUREMENT_ROLES and not self._notif_started:
            self._notif_started = True
            self._start_notifications()

    def _make_settings_bar(self, parent, reload_fn):
        C = T()
        bar = ctk.CTkFrame(parent, fg_color=C['surface2'], corner_radius=0, height=40)
        bar.pack(fill='x', side='top')
        
        right = ctk.CTkFrame(bar, fg_color=C['surface2']); right.pack(side='right', padx=16)
        
        def toggle_theme():
            SETTINGS['theme'] = 'light' if SETTINGS.get('theme','dark')=='dark' else 'dark'
            save_settings(SETTINGS)
            reload_fn()
        
        is_dark = SETTINGS.get('theme','dark') == 'dark'
        icon = '☀  Light Mode' if is_dark else '🌙  Dark Mode'
        ghost_btn(right, icon, toggle_theme, color=C['gold'], w=140, h=28).pack(side='left', pady=6, padx=4)

    # ══ TOPBAR ═══════════════════════════════════════════════
    def _topbar(self, parent, subtitle, title, reload_fn):
        C = T()
        wrap = ctk.CTkFrame(parent, fg_color=C['bg']); wrap.pack(fill='x')
        
        # Settings strip
        strip = ctk.CTkFrame(wrap, fg_color=C['surface2'], corner_radius=0, height=36)
        strip.pack(fill='x'); strip.pack_propagate(False)
        right_strip = ctk.CTkFrame(strip, fg_color=C['surface2']); right_strip.pack(side='right', padx=12)
        
        def toggle_theme():
            SETTINGS['theme'] = 'light' if SETTINGS.get('theme','dark')=='dark' else 'dark'
            save_settings(SETTINGS)
            reload_fn()
        
        is_dark = SETTINGS.get('theme','dark') == 'dark'
        icon = '☀  Light' if is_dark else '🌙  Dark'
        ghost_btn(right_strip, icon, toggle_theme, color=C['gold'], w=100, h=26).pack(side='left', pady=5, padx=4)
        ghost_btn(right_strip, '⊞  Modes', self._modes, w=90, h=26).pack(side='left', pady=5, padx=4)

        # Main topbar
        bar = ctk.CTkFrame(wrap, fg_color=C['topbar'], corner_radius=0, height=80)
        bar.pack(fill='x'); bar.pack_propagate(False)
        ctk.CTkFrame(bar, height=1, fg_color=C['border']).pack(side='bottom', fill='x')
        
        # Logo
        _timg = load_logo('topbar', 'bsb_logo_topbar.png', (36,40))
        if _timg:
            tk.Label(bar, image=_timg,
                    bg=C['topbar'], bd=0, highlightthickness=0).pack(side='left', padx=(16,8), pady=10)
        
        left = ctk.CTkFrame(bar, fg_color=C['topbar']); left.pack(side='left', pady=10)
        L(left, subtitle, C['text4'], 'xs').pack(anchor='w')
        L(left, title, C['text'], 'h2').pack(anchor='w')
        
        right = ctk.CTkFrame(bar, fg_color=C['topbar']); right.pack(side='right', padx=16)
        badge = ctk.CTkFrame(right, fg_color=C['surface2'], corner_radius=8,
                            border_width=1, border_color=C['border'])
        badge.pack(side='left', padx=8)
        av = ctk.CTkFrame(badge, width=28, height=28, corner_radius=14, fg_color=C['gold'])
        av.pack(side='left', padx=(8,7), pady=7); av.pack_propagate(False)
        ctk.CTkLabel(av, text=self.current_user[0].upper(), text_color='#0B1F3A',
                    font=ctk.CTkFont('Outfit',13,'bold')).place(relx=0.5,rely=0.5,anchor='center')
        col = ctk.CTkFrame(badge, fg_color=C['surface2']); col.pack(side='left', padx=(0,12), pady=5)
        L(col, self.current_user, C['text'], 'sm_b').pack(anchor='w')
        L(col, self.current_role.upper(), C['gold'], 'caps').pack(anchor='w')
        red_btn(right, 'Sign Out', self._login, w=100, h=34).pack(side='left')

    # ── Shared sidebar ────────────────────────────────────────
    def _sidebar(self, parent, on_filter):
        C = T()
        sb = tk.Frame(parent, bg=C['sidebar'], width=240); sb.pack(side='left', fill='y')
        sb.pack_propagate(False)
        tk.Frame(sb, bg='#1A3A5C', width=1).pack(side='right', fill='y')

        # Logo + title
        hdr = tk.Frame(sb, bg=C['sidebar']); hdr.pack(fill='x', padx=14, pady=(14,12))
        _simg = load_logo('sidebar', 'bsb_logo_sidebar.png', (44,50))
        if _simg:
            tk.Label(hdr, image=_simg, bg=C['sidebar'], bd=0).pack(side='left', padx=(0,10))
        else:
            lf = tk.Frame(hdr, bg=C['gold'], width=40, height=40); lf.pack(side='left', padx=(0,10))
            lf.pack_propagate(False)
            tk.Label(lf, text='BSB', bg=C['gold'], fg='#0B1F3A', font=('Cormorant Garamond',14,'bold')).place(relx=0.5,rely=0.5,anchor='center')
        txt = tk.Frame(hdr, bg=C['sidebar']); txt.pack(side='left')
        tk.Label(txt, text='Stationery', bg=C['sidebar'], fg=C['side_txt'], font=('Outfit',13,'bold')).pack(anchor='w')
        tk.Label(txt, text='British School of Bahrain', bg=C['sidebar'], fg=C['side_mut'], font=('Outfit',10)).pack(anchor='w')
        tk.Frame(sb, bg='#132944', height=1).pack(fill='x')

        # CATEGORIES label
        tk.Label(sb, text='CATEGORIES', bg=C['sidebar'], fg=C['side_mut'],
                font=('Outfit',9,'bold'), anchor='w', padx=16, pady=12).pack(fill='x')

        # Stock alerts — ONLY admin/procurement. Staff must NEVER see stock levels.
        show_stock = (self.current_role in PROCUREMENT_ROLES) and (self.current_role != 'staff')
        if show_stock:
            tk.Frame(sb, bg='#132944', height=1).pack(fill='x', side='bottom')
            alerts = tk.Frame(sb, bg=C['sidebar']); alerts.pack(side='bottom', fill='x', padx=12, pady=10)
            stats = qy("SELECT SUM(CASE WHEN Quantity=0 THEN 1 ELSE 0 END), SUM(CASE WHEN Quantity>0 AND Quantity<MinStock THEN 1 ELSE 0 END) FROM Items")
            out_n = stats[0][0] or 0; low_n = stats[0][1] or 0
            tk.Label(alerts, text='STOCK ALERTS', bg=C['sidebar'], fg=C['side_mut'],
                    font=('Outfit',9,'bold'), anchor='w').pack(fill='x', pady=(0,6))
            row2 = tk.Frame(alerts, bg=C['sidebar']); row2.pack(fill='x')
            for val, label, clr in [(low_n,'Low',C['orange']),(out_n,'Out',C['red'])]:
                col = tk.Frame(row2, bg=C['sidebar']); col.pack(side='left', expand=True)
                tk.Label(col, text=str(val), bg=C['sidebar'], fg=clr, font=('Outfit',20,'bold')).pack(anchor='w')
                tk.Label(col, text=label, bg=C['sidebar'], fg=C['side_mut'], font=('Outfit',10)).pack(anchor='w')

        # Scrollable nav
        canvas = tk.Canvas(sb, bg=C['sidebar'], highlightthickness=0, bd=0)
        scr = tk.Scrollbar(sb, orient='vertical', command=canvas.yview)
        nav = tk.Frame(canvas, bg=C['sidebar'])
        nav_id = canvas.create_window((0,0), window=nav, anchor='nw', width=222)
        canvas.configure(yscrollcommand=scr.set)
        canvas.pack(side='left', fill='both', expand=True)
        scr.pack(side='right', fill='y')
        def _on_nav_config(e):
            canvas.configure(scrollregion=canvas.bbox('all'))
        nav.bind('<Configure>', _on_nav_config)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(nav_id, width=e.width))
        # Mouse wheel scrolling
        def _wheel(e):
            canvas.yview_scroll(int(-1*(e.delta/120)), 'units')
        def _bind_wheel(_):  canvas.bind_all('<MouseWheel>', _wheel)
        def _unbind_wheel(_): canvas.unbind_all('<MouseWheel>')
        canvas.bind('<Enter>', _bind_wheel)
        canvas.bind('<Leave>', _unbind_wheel)
        sb.bind('<Enter>', _bind_wheel)

        active = [None]
        def set_active(b):
            if active[0]:
                try: active[0].configure(bg=C['sidebar'], fg=C['side_mut'])
                except: pass
            active[0] = b; b.configure(bg='#132944', fg=C['gold'])

        all_row = tk.Frame(nav, bg='#132944'); all_row.pack(fill='x')
        all_btn = tk.Button(all_row, text='  ⊞  All Items', bg='#132944', fg=C['gold'],
                           font=('Outfit',13,'bold'), relief='flat', anchor='w',
                           padx=12, pady=10, cursor='hand2', bd=0, highlightthickness=0,
                           activebackground='#132944', activeforeground=C['gold'])
        all_btn.pack(side='left', fill='x', expand=True)
        total_items = (qy("SELECT COUNT(*) FROM Items") or [(0,)])[0][0]
        tk.Label(all_row, text=f' {total_items} ', bg=C['blue'], fg='white',
                font=('Outfit',9,'bold')).pack(side='right', padx=8)
        all_btn.configure(command=lambda: (set_active(all_btn), on_filter(None,None)))
        active[0] = all_btn

        for pr in qy("SELECT ParentCatID, Name FROM ParentCategories ORDER BY Name"):
            pid, pname = pr[0], pr[1]
            subs = qy("SELECT SubCatID, Name FROM SubCategories WHERE ParentCatID=? ORDER BY Name",(pid,))
            if not subs:
                continue  # don't show an empty parent heading
            grp = tk.Label(nav, text=pname.upper(), bg=C['sidebar'], fg=C['side_mut'],
                          font=('Outfit',9,'bold'), anchor='w', padx=16, pady=8)
            grp.pack(fill='x')
            for sr in subs:
                sid, sname = sr[0], sr[1]
                low = (qy("SELECT COUNT(*) FROM Items WHERE SubCatID=? AND Quantity<MinStock AND Quantity>0",(sid,)) or [(0,)])[0][0]
                row = tk.Frame(nav, bg=C['sidebar']); row.pack(fill='x')
                sbtn = tk.Button(row, text=f'  {sname}', bg=C['sidebar'], fg=C['side_mut'],
                                font=('Outfit',12,'bold'), relief='flat', anchor='w',
                                padx=12, pady=8, cursor='hand2', bd=0, highlightthickness=0,
                                activebackground='#132944', activeforeground=C['gold'],
                                wraplength=185, justify='left')
                sbtn.pack(side='left', fill='x', expand=True)
                cnt = (qy("SELECT COUNT(*) FROM Items WHERE SubCatID=?",(sid,)) or [(0,)])[0][0]
                if low > 0 and show_stock:
                    tk.Label(row, text='●', bg=C['sidebar'], fg=C['orange'],
                            font=('Outfit',9)).pack(side='right')
                tk.Label(row, text=f' {cnt} ', bg='#132944', fg=C['side_mut'],
                        font=('Outfit',9,'bold')).pack(side='right', padx=6)
                sbtn.configure(command=lambda s=sid, p=pid, b=sbtn: (set_active(b), on_filter(p,s)))

        # ── Orphan subcategories (NULL / missing parent) — show so nothing is hidden ──
        orphans = qy("""SELECT s.SubCatID, s.Name FROM SubCategories s
                        WHERE s.ParentCatID IS NULL
                           OR s.ParentCatID NOT IN (SELECT ParentCatID FROM ParentCategories)
                        ORDER BY s.Name""")
        if orphans:
            grp = tk.Label(nav, text='OTHER', bg=C['sidebar'], fg=C['side_mut'],
                          font=('Outfit',9,'bold'), anchor='w', padx=16, pady=8)
            grp.pack(fill='x')
            for sr in orphans:
                sid, sname = sr[0], sr[1]
                row = tk.Frame(nav, bg=C['sidebar']); row.pack(fill='x')
                sbtn = tk.Button(row, text=f'  {sname}', bg=C['sidebar'], fg=C['side_mut'],
                                font=('Outfit',12,'bold'), relief='flat', anchor='w',
                                padx=12, pady=8, cursor='hand2', bd=0, highlightthickness=0,
                                activebackground='#132944', activeforeground=C['gold'],
                                wraplength=185, justify='left')
                sbtn.pack(side='left', fill='x', expand=True)
                cnt = (qy("SELECT COUNT(*) FROM Items WHERE SubCatID=?",(sid,)) or [(0,)])[0][0]
                tk.Label(row, text=f' {cnt} ', bg='#132944', fg=C['side_mut'],
                        font=('Outfit',9,'bold')).pack(side='right', padx=6)
                sbtn.configure(command=lambda s=sid, b=sbtn: (set_active(b), on_filter(None,s)))

    # ══ MANAGE ═══════════════════════════════════════════════
    def _manage(self):
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root = ctk.CTkFrame(self, fg_color=C['bg']); root.pack(fill='both', expand=True)
        self._topbar(root, 'Inventory', 'Manage Inventory', lambda: self._reload(self._manage))

        # Tab bar
        tabbar = ctk.CTkFrame(root, fg_color=C['topbar'], corner_radius=0, height=48)
        tabbar.pack(fill='x'); tabbar.pack_propagate(False)
        hline(tabbar, 0)
        content = ctk.CTkFrame(root, fg_color=C['bg']); content.pack(fill='both', expand=True)

        tabs = []
        def switch(name, fn):
            for w in content.winfo_children(): w.destroy()
            fn(content)
            for b, n, ul in tabs:
                on = (n==name)
                b.configure(fg_color=C['topbar'], text_color=C['text'] if on else C['text4'], hover_color=C['surface2'])
                ul.configure(fg_color=C['gold'] if on else C['topbar'])
        for label, name, fn in [('📦  Stock Items','stock',self._tab_stock),('🏷  Categories','cats',self._tab_cats)]:
            wrap = tk.Frame(tabbar, bg=C['topbar']); wrap.pack(side='left')
            b = ctk.CTkButton(wrap, text=label, fg_color=C['topbar'], hover_color=C['surface2'],
                             text_color=C['text'] if name=='stock' else C['text4'],
                             font=F['sm_b'], corner_radius=0, height=48, width=175,
                             command=lambda n=name, f=fn: switch(n,f))
            b.pack()
            ul = ctk.CTkFrame(wrap, height=3, fg_color=C['gold'] if name=='stock' else C['topbar'])
            ul.pack(fill='x'); tabs.append((b,name,ul))
        self._tab_stock(content)

    # ══════════════════════════════════════════════════════════
    # IT ASSET MANAGEMENT — completely separate from Procurement.
    # Uses ITAssets + ITCategories tables. Never touches Items.
    # ══════════════════════════════════════════════════════════
    def _it_manage(self, decom_mode=False):
        """IT Manage entry point: a searchable grid of category tiles.
        decom_mode=True shows only the decommission categories."""
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        self._make_settings_bar(self, lambda: self._reload(lambda: self._it_manage(decom_mode)))

        outer = ctk.CTkScrollableFrame(self, fg_color=C['bg'])
        outer.pack(fill='both', expand=True)
        center = ctk.CTkFrame(outer, fg_color=C['bg']); center.pack(pady=(24,24))

        L(center, 'Decommission' if decom_mode else 'IT Manage', C['text'], 'disp').pack(pady=(0,4))
        L(center, 'PICK A CATEGORY', C['gold'], 'caps').pack(pady=(0,10))
        L(center, ('Old / retired devices kept for records.' if decom_mode
                   else 'Search or click a category to view, add, or edit its assets.'),
          C['text3'], 'sm').pack(pady=(0,16))

        # ── Search bar to filter the category tiles ──
        search_var = ctk.StringVar()
        srow = ctk.CTkFrame(center, fg_color=C['bg']); srow.pack(pady=(0,22))
        ctk.CTkEntry(srow, textvariable=search_var, placeholder_text='🔎  Search categories...',
                     width=420, height=42, fg_color=C['surface2'], border_color=C['border'],
                     text_color=C['text'], placeholder_text_color=C['text4'],
                     font=F['body'], corner_radius=8).pack()

        grid_holder = ctk.CTkFrame(center, fg_color=C['bg']); grid_holder.pack()

        cats = qy("SELECT ITCatID, Name FROM ITCategories ORDER BY Name")
        total = (qy("SELECT COUNT(*) FROM ITAssets") or [(0,)])[0][0]

        # Split categories into normal vs decom
        decom_set = set(IT_DECOM_CATEGORIES)
        def in_scope(name):
            return (name in decom_set) if decom_mode else (name not in decom_set)

        def build_grid(*_):
            for w in grid_holder.winfo_children(): w.destroy()
            q = search_var.get().strip().lower()
            tiles = []
            if not decom_mode and (not q or 'all' in q):
                tiles.append(('🗂', 'All Assets', f'{total} total',
                              lambda: self._it_category_screen(None, None)))
            for cid, cname in cats:
                if not in_scope(cname): continue
                if q and q not in cname.lower(): continue
                cnt = (qy("SELECT COUNT(*) FROM ITAssets WHERE ITCatID=?",(cid,)) or [(0,)])[0][0]
                icon = '🗑' if cname in decom_set else '💻'
                tiles.append((icon, cname, f'{cnt} asset{"s" if cnt!=1 else ""}',
                              lambda c=cid, n=cname: self._it_category_screen(c, n)))
            if not tiles:
                L(grid_holder, 'No categories match your search.', C['text3'], 'sm').pack(pady=20)
                return
            cur_row = None
            for idx, (icon, name, desc, cmd) in enumerate(tiles):
                if idx % 4 == 0:
                    cur_row = ctk.CTkFrame(grid_holder, fg_color=C['bg']); cur_row.pack(pady=8)
                card = ctk.CTkFrame(cur_row, fg_color=C['surface'], corner_radius=10,
                                   border_width=1, border_color=C['border'], width=190, height=130)
                card.pack(side='left', padx=10); card.pack_propagate(False)
                L(card, icon, fk='h1').pack(pady=(16,4))
                L(card, name, C['text'], 'body_b').pack(padx=10)
                L(card, desc, C['text3'], 'tiny').pack(pady=(4,0))
                for w in [card]+list(card.winfo_children()):
                    w.bind('<Button-1>', lambda e, c2=cmd: c2())
                    w.bind('<Enter>', lambda e, w=card: w.configure(border_color=C['gold']))
                    w.bind('<Leave>', lambda e, w=card: w.configure(border_color=C['border']))

        search_var.trace_add('write', build_grid)
        build_grid()

        hline(center, 24)
        bbar = ctk.CTkFrame(center, fg_color=C['bg']); bbar.pack(pady=(0,20))
        if not decom_mode:
            gold_btn(bbar, '⬆ Bulk Upload (all sheets)',
                     lambda: self._it_bulk(lambda: self._reload(self._it_manage)),
                     w=210, h=40).pack(side='left', padx=6)
            ghost_btn(bbar, '🏷 Manage Categories', self._it_cats_screen,
                      w=190, h=40).pack(side='left', padx=6)
            ghost_btn(bbar, '🗑 Decommission', lambda: self._it_manage(decom_mode=True),
                      w=160, h=40).pack(side='left', padx=6)
            back_fn = self._dept_it if self.current_role == 'admin' else self._modes
            ghost_btn(bbar, '⬅  IT Modes', back_fn, w=130, h=40).pack(side='left', padx=6)
        else:
            ghost_btn(bbar, '⬅  Back to Manage', lambda: self._reload(self._it_manage),
                      w=180, h=40).pack(side='left', padx=6)

    def _it_cats_screen(self):
        """Standalone categories management screen (add/delete IT categories)."""
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root = ctk.CTkFrame(self, fg_color=C['bg']); root.pack(fill='both', expand=True)
        self._topbar(root, 'IT Department', 'Manage Categories', lambda: self._reload(self._it_cats_screen))
        content = ctk.CTkFrame(root, fg_color=C['bg']); content.pack(fill='both', expand=True)
        self._it_tab_cats(content)
        bbar = ctk.CTkFrame(root, fg_color=C['bg'], height=50); bbar.pack(fill='x'); bbar.pack_propagate(False)
        ghost_btn(bbar, '⬅  Back to Categories', lambda: self._reload(self._it_manage), w=180, h=36).pack(side='left', padx=18, pady=8)

    def _it_category_screen(self, cat_id, cat_name):
        """Category screen with a smart left sidebar.
        The sidebar shows the DISTINCT VALUES of the most useful grouping
        field for THAT category (e.g. Location for Network Devices, Building
        for Desktops, Device Status for Kiosk Machines).
        Clicking a value filters the table to just those assets."""
        import json
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root = ctk.CTkFrame(self, fg_color=C['bg']); root.pack(fill='both', expand=True)
        title = cat_name if cat_name else 'All Assets'
        self._topbar(root, 'IT Department', title,
                     lambda: self._reload(lambda: self._it_category_screen(cat_id, cat_name)))

        body = ctk.CTkFrame(root, fg_color=C['bg']); body.pack(fill='both', expand=True)

        # ── Left sidebar ──────────────────────────────────────────
        # For a specific category: shows distinct values of its key field.
        # For "All Assets" (cat_id None): no sidebar — full width table.
        sidebar_field = IT_SIDEBAR_FIELD.get(cat_name) if cat_name else None
        filt = {'value': None}  # currently selected sidebar value (None = All)
        active_btn = [None]

        if sidebar_field and cat_id is not None:
            sb = tk.Frame(body, bg=C['sidebar'], width=220)
            sb.pack(side='left', fill='y'); sb.pack_propagate(False)
            tk.Frame(sb, bg='#1A3A5C', width=1).pack(side='right', fill='y')

            # Header
            tk.Label(sb, text=cat_name.upper(), bg=C['sidebar'], fg=C['gold'],
                     font=('Outfit',9,'bold'), anchor='w', padx=14, pady=12).pack(fill='x')
            tk.Label(sb, text=sidebar_field.upper(), bg=C['sidebar'], fg=C['side_mut'],
                     font=('Outfit',8,'bold'), anchor='w', padx=14).pack(fill='x', pady=(0,6))
            tk.Frame(sb, bg='#132944', height=1).pack(fill='x')

            # Scrollable filter list
            canvas2 = tk.Canvas(sb, bg=C['sidebar'], highlightthickness=0, bd=0)
            scr2 = tk.Scrollbar(sb, orient='vertical', command=canvas2.yview)
            nav2 = tk.Frame(canvas2, bg=C['sidebar'])
            nid2 = canvas2.create_window((0,0), window=nav2, anchor='nw', width=202)
            canvas2.configure(yscrollcommand=scr2.set)
            canvas2.pack(side='left', fill='both', expand=True)
            scr2.pack(side='right', fill='y')
            nav2.bind('<Configure>', lambda e: canvas2.configure(scrollregion=canvas2.bbox('all')))
            canvas2.bind('<Configure>', lambda e: canvas2.itemconfig(nid2, width=e.width))
            def _w2(e): canvas2.yview_scroll(int(-1*(e.delta/120)), 'units')
            canvas2.bind('<Enter>', lambda _: canvas2.bind_all('<MouseWheel>', _w2))
            canvas2.bind('<Leave>', lambda _: canvas2.unbind_all('<MouseWheel>'))

            # Back button at bottom
            tk.Frame(sb, bg='#132944', height=1).pack(side='bottom', fill='x')
            tk.Button(sb, text='⬅  Back to Manage', bg=C['sidebar'], fg=C['side_mut'],
                      font=('Outfit',10,'bold'), relief='flat', anchor='w', padx=14, pady=9,
                      cursor='hand2', bd=0, highlightthickness=0,
                      activebackground='#132944', activeforeground=C['text'],
                      command=lambda: self._reload(self._it_manage)).pack(side='bottom', fill='x')

            def set_sb_active(btn):
                if active_btn[0]:
                    active_btn[0].configure(bg=C['sidebar'], fg=C['side_mut'])
                btn.configure(bg='#132944', fg=C['gold'])
                active_btn[0] = btn

            # Compute the distinct values + counts ONCE (not per-button)
            raw = qy("SELECT ISNULL(DetailsJSON,'') FROM ITAssets WHERE ITCatID=?",(cat_id,))
            seen = {}; total_here = 0
            for row in raw:
                djson = row[0] if isinstance(row,(tuple,list)) else row
                try: d = json.loads(djson) if djson else {}
                except: d = {}
                total_here += 1
                v = d.get(sidebar_field,'').strip()
                if v:
                    seen[v] = seen.get(v,0) + 1

            def make_filter_btn(val_label, filter_val, cnt):
                """val_label = display text, filter_val = None means All"""
                row = tk.Frame(nav2, bg=C['sidebar']); row.pack(fill='x')
                btn = tk.Button(row, text=f'  {val_label}',
                                bg=C['sidebar'], fg=C['side_mut'],
                                font=('Outfit',11,'bold'), relief='flat', anchor='w',
                                padx=10, pady=7, cursor='hand2', bd=0, highlightthickness=0,
                                activebackground='#132944', activeforeground=C['gold'],
                                wraplength=170, justify='left')
                btn.pack(side='left', fill='x', expand=True)
                tk.Label(row, text=f' {cnt} ', bg='#132944', fg=C['side_mut'],
                         font=('Outfit',9,'bold')).pack(side='right', padx=5)
                def on_click(fv=filter_val, b=btn):
                    filt['value'] = fv
                    set_sb_active(b)
                    load()
                btn.configure(command=on_click)
                return btn

            # "All" button first
            all_sb_btn = make_filter_btn('⊞  All', None, total_here)
            set_sb_active(all_sb_btn)  # default active
            for val in sorted(seen.keys()):
                make_filter_btn(val, val, seen[val])

        else:
            # No sidebar for All Assets or categories without a key field
            # Back button as a simple frame instead
            pass

        # ── Right content area ────────────────────────────────────
        right = ctk.CTkFrame(body, fg_color=C['bg'])
        right.pack(side='left', fill='both', expand=True)

        # Determine the table columns for this category
        if cat_name and cat_name in IT_CATEGORY_FIELDS:
            data_fields = list(IT_CATEGORY_FIELDS[cat_name])
            cols = tuple(data_fields[:8]) + ('Stock',)
        else:
            data_fields = ['Name','Category','Model','Serial','Location']
            cols = ('Name','Category','Model','Serial','Location','Stock')

        # Search-column choices: "All Columns" + each data field
        search_cols = ['All Columns'] + list(cols[:-1])  # exclude Stock from search picker
        search_col_var = ctk.StringVar(value='All Columns')

        # Toolbar
        sv = ctk.StringVar()
        tb = ctk.CTkFrame(right, fg_color=C['topbar'], corner_radius=0, height=56)
        tb.pack(fill='x'); tb.pack_propagate(False)
        L(tb,'Search',C['text4'],'xs').pack(side='left',padx=(14,6),pady=14)
        ctk.CTkEntry(tb, textvariable=sv, placeholder_text='Search...', width=200, height=34,
                     fg_color=C['surface2'], border_color=C['border'], text_color=C['text'],
                     placeholder_text_color=C['text4'], font=F['body'], corner_radius=6
                     ).pack(side='left',pady=12)
        # Filter dropdown: which column to search
        ctk.CTkComboBox(tb, values=search_cols, variable=search_col_var, width=150, height=34,
                        fg_color=C['surface2'], border_color=C['border'], button_color=C['gold'],
                        button_hover_color=C['gold2'], text_color=C['text'], font=F['sm'],
                        corner_radius=6, dropdown_fg_color=C['surface2'], dropdown_text_color=C['text'],
                        command=lambda *_: load()).pack(side='left', padx=8, pady=12)

        blue_btn(tb,'＋ Add',lambda: self._it_asset_dlg_for(cat_id,cat_name,reload),w=90,h=34).pack(side='right',padx=4,pady=10)
        if cat_id is not None:
            gold_btn(tb,'⬆ Bulk',lambda: self._it_bulk_one(cat_name,reload),w=100,h=34).pack(side='right',padx=4,pady=10)
        ghost_btn(tb,'✎ Edit',lambda: self._it_asset_dlg_for(cat_id,cat_name,reload,True),w=70,h=34).pack(side='right',padx=4,pady=10)
        red_btn(tb,'🗑 Del',lambda: self._it_del(reload),w=70,h=34).pack(side='right',padx=4,pady=10)

        # Back button in toolbar row if no sidebar
        if not (sidebar_field and cat_id is not None):
            ghost_btn(tb,'⬅ Back',lambda: self._reload(self._it_manage),w=80,h=34).pack(side='left',padx=8,pady=10)

        hint = ctk.CTkFrame(right, fg_color=C['bg'], height=24); hint.pack(fill='x'); hint.pack_propagate(False)
        hint_txt = f'Filtering by {sidebar_field}  ·  ' if sidebar_field and cat_id else ''
        L(hint, f'{hint_txt}Pick a column in the dropdown to search just that field  ·  ☐ in header = select all',
          C['text4'], 'tiny').pack(side='left',padx=14,pady=3)

        tcard = ctk.CTkFrame(right, fg_color=C['card'], corner_radius=10,
                             border_width=1, border_color=C['border'])
        tcard.pack(fill='both', expand=True, padx=10, pady=(0,10))

        widths = [max(80, 920//len(cols))]*len(cols); widths[-1] = 60
        cbt = CheckboxTree(tcard, cols, widths, height=18, simple_mode=True)
        cbt.frame.pack(fill='both', expand=True, padx=2, pady=2)
        self._it_tree = cbt
        cbt.tree.bind('<Double-1>', lambda e: self._it_asset_dlg_for(cat_id,cat_name,reload,True))

        status_field = IT_STATUS_FIELD.get(cat_name)

        def row_status_tag(details):
            """Return 'status_red'/'status_amber'/'' based on the category's status field."""
            if not status_field: return ''
            val = (details.get(status_field,'') or '').strip().lower()
            if not val: return ''
            if any(k in val for k in IT_STATUS_RED): return 'status_red'
            if any(k in val for k in IT_STATUS_AMBER): return 'status_amber'
            return ''

        def load():
            cbt.clear()
            search_text = sv.get().strip().lower()
            search_col = search_col_var.get()
            try:
                if cat_id is not None:
                    rows = qy("""SELECT a.AssetID, ISNULL(a.DetailsJSON,''), a.Stock
                                 FROM ITAssets a WHERE a.ITCatID=? ORDER BY a.AssetID""",(cat_id,))
                else:
                    rows = qy("""SELECT a.AssetID, ISNULL(a.DetailsJSON,''), a.Stock, ISNULL(c.Name,'-')
                                 FROM ITAssets a LEFT JOIN ITCategories c ON a.ITCatID=c.ITCatID
                                 ORDER BY c.Name, a.AssetID""")
            except Exception as e:
                messagebox.showerror('Load error', str(e), parent=self); return

            for r in rows:
                if cat_id is not None:
                    aid, djson, stock = r
                    details = {}
                    if djson:
                        try: details = json.loads(djson)
                        except: pass
                    # sidebar filter
                    if filt['value'] is not None and details.get(sidebar_field,'').strip() != filt['value']:
                        continue
                    field_vals = [details.get(f,'') for f in data_fields[:8]]
                    vals = tuple(field_vals) + (stock,)
                    # search filter
                    if search_text:
                        if search_col == 'All Columns':
                            hay = ' '.join(str(v).lower() for v in vals)
                        else:
                            try: idx = list(cols).index(search_col)
                            except ValueError: idx = 0
                            hay = str(vals[idx]).lower() if idx < len(vals) else ''
                        if search_text not in hay:
                            continue
                    tag = row_status_tag(details)
                    cbt.insert(aid, vals, tag)
                else:
                    aid, djson, stock, catn = r
                    details = {}
                    if djson:
                        try: details = json.loads(djson)
                        except: pass
                    name = details.get(list(details.keys())[0],'') if details else ''
                    # best-effort generic columns
                    name = (details.get('Asset Name') or details.get('Asset Tag') or
                            details.get('Device Name') or details.get('Name') or
                            details.get('Printer Name') or details.get('Brand') or
                            (list(details.values())[0] if details else ''))
                    model = details.get('Model') or details.get('Model No.') or ''
                    serial = details.get('Serial') or details.get('S/N') or details.get('Serial Number') or ''
                    loc = details.get('Location') or details.get('Class/Office') or details.get('Default Location') or ''
                    vals = (name, catn, model, serial, loc, stock)
                    if search_text:
                        if search_col == 'All Columns':
                            hay = ' '.join(str(v).lower() for v in vals)
                        else:
                            try: idx = list(cols).index(search_col)
                            except ValueError: idx = 0
                            hay = str(vals[idx]).lower() if idx < len(vals) else ''
                        if search_text not in hay: continue
                    cbt.insert(aid, vals)

        def reload():
            self._reload(lambda: self._it_category_screen(cat_id, cat_name))

        sv.trace_add('write', lambda *_: load())
        load()

    # ── IT sidebar (categories from ITCategories) ─────────────
    def _it_sidebar(self, parent, on_filter):
        C = T()
        sb = tk.Frame(parent, bg=C['sidebar'], width=270); sb.pack(side='left', fill='y')
        sb.pack_propagate(False)
        tk.Frame(sb, bg='#1A3A5C', width=1).pack(side='right', fill='y')

        hdr = tk.Frame(sb, bg=C['sidebar']); hdr.pack(fill='x', padx=14, pady=(14,12))
        _simg = load_logo('sidebar', 'bsb_logo_sidebar.png', (44,50))
        if _simg:
            tk.Label(hdr, image=_simg, bg=C['sidebar'], bd=0).pack(side='left', padx=(0,10))
        else:
            lf = tk.Frame(hdr, bg=C['gold'], width=40, height=40); lf.pack(side='left', padx=(0,10)); lf.pack_propagate(False)
            tk.Label(lf, text='BSB', bg=C['gold'], fg='#0B1F3A', font=('Cormorant Garamond',14,'bold')).place(relx=0.5,rely=0.5,anchor='center')
        txt = tk.Frame(hdr, bg=C['sidebar']); txt.pack(side='left')
        tk.Label(txt, text='IT Assets', bg=C['sidebar'], fg=C['side_txt'], font=('Outfit',13,'bold')).pack(anchor='w')
        tk.Label(txt, text='British School of Bahrain', bg=C['sidebar'], fg=C['side_mut'], font=('Outfit',10)).pack(anchor='w')
        tk.Frame(sb, bg='#132944', height=1).pack(fill='x')

        tk.Label(sb, text='CATEGORIES', bg=C['sidebar'], fg=C['side_mut'],
                font=('Outfit',9,'bold'), anchor='w', padx=16, pady=12).pack(fill='x')

        canvas = tk.Canvas(sb, bg=C['sidebar'], highlightthickness=0, bd=0)
        scr = tk.Scrollbar(sb, orient='vertical', command=canvas.yview)
        nav = tk.Frame(canvas, bg=C['sidebar'])
        nav_id = canvas.create_window((0,0), window=nav, anchor='nw', width=252)
        canvas.configure(yscrollcommand=scr.set)
        canvas.pack(side='left', fill='both', expand=True)
        scr.pack(side='right', fill='y')
        nav.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(nav_id, width=e.width))
        def _wheel(e): canvas.yview_scroll(int(-1*(e.delta/120)), 'units')
        canvas.bind('<Enter>', lambda _: canvas.bind_all('<MouseWheel>', _wheel))
        canvas.bind('<Leave>', lambda _: canvas.unbind_all('<MouseWheel>'))

        active = [None]
        def set_active(b):
            if active[0]:
                try: active[0].configure(bg=C['sidebar'], fg=C['side_mut'])
                except: pass
            active[0] = b; b.configure(bg='#132944', fg=C['gold'])

        all_row = tk.Frame(nav, bg='#132944'); all_row.pack(fill='x')
        all_btn = tk.Button(all_row, text='  ⊞  All Assets', bg='#132944', fg=C['gold'],
                           font=('Outfit',13,'bold'), relief='flat', anchor='w',
                           padx=12, pady=10, cursor='hand2', bd=0, highlightthickness=0,
                           activebackground='#132944', activeforeground=C['gold'])
        all_btn.pack(side='left', fill='x', expand=True)
        total = (qy("SELECT COUNT(*) FROM ITAssets") or [(0,)])[0][0]
        tk.Label(all_row, text=f' {total} ', bg=C['blue'], fg='white', font=('Outfit',9,'bold')).pack(side='right', padx=8)
        all_btn.configure(command=lambda: (set_active(all_btn), on_filter(None, None)))
        active[0] = all_btn

        for cr in qy("SELECT ITCatID, Name FROM ITCategories ORDER BY Name"):
            cid, cname = cr[0], cr[1]
            row = tk.Frame(nav, bg=C['sidebar']); row.pack(fill='x')
            cbtn = tk.Button(row, text=f'  {cname}', bg=C['sidebar'], fg=C['side_mut'],
                            font=('Outfit',12,'bold'), relief='flat', anchor='w',
                            padx=12, pady=8, cursor='hand2', bd=0, highlightthickness=0,
                            activebackground='#132944', activeforeground=C['gold'],
                            wraplength=185, justify='left')
            cbtn.pack(side='left', fill='x', expand=True)
            cnt = (qy("SELECT COUNT(*) FROM ITAssets WHERE ITCatID=?",(cid,)) or [(0,)])[0][0]
            tk.Label(row, text=f' {cnt} ', bg='#132944', fg=C['side_mut'], font=('Outfit',9,'bold')).pack(side='right', padx=6)
            cbtn.configure(command=lambda c=cid, nm=cname, b=cbtn: (set_active(b), on_filter(c, nm)))

    # ── IT assets tab ─────────────────────────────────────────
    def _it_tab_assets(self, parent):
        import json
        C = T()
        wrap = ctk.CTkFrame(parent, fg_color=C['bg']); wrap.pack(fill='both', expand=True)
        filt = {'c':None, 'cname':None}
        def on_filter(c, cname=None):
            filt['c']=c; filt['cname']=cname; rebuild_table()
        self._it_sidebar(wrap, on_filter)

        right = ctk.CTkFrame(wrap, fg_color=C['bg']); right.pack(side='left', fill='both', expand=True)

        tb = ctk.CTkFrame(right, fg_color=C['topbar'], corner_radius=0, height=60)
        tb.pack(fill='x'); tb.pack_propagate(False)
        sv = ctk.StringVar()
        L(tb,'Search',C['text4'],'xs').pack(side='left',padx=(18,8),pady=16)
        ctk.CTkEntry(tb, textvariable=sv, placeholder_text='Search assets...',
                    width=260, height=36, fg_color=C['surface2'], border_color=C['border'],
                    text_color=C['text'], placeholder_text_color=C['text4'],
                    font=F['body'], corner_radius=6).pack(side='left',pady=14)
        blue_btn(tb,'＋ Add Asset',lambda: self._it_asset_dlg(reload),w=120,h=36).pack(side='right',padx=4,pady=12)
        gold_btn(tb,'⬆ Bulk Upload',lambda: self._it_bulk(reload),w=130,h=36).pack(side='right',padx=4,pady=12)
        ghost_btn(tb,'✎ Edit',lambda: self._it_asset_dlg(reload,True),w=80,h=36).pack(side='right',padx=4,pady=12)
        red_btn(tb,'🗑 Delete',lambda: self._it_del(reload),w=90,h=36).pack(side='right',padx=4,pady=12)

        hint = ctk.CTkFrame(right, fg_color=C['bg'], height=26); hint.pack(fill='x'); hint.pack_propagate(False)
        L(hint,'ⓘ  Check items to select (☐ in the header selects/deselects all) · Pick a category on the left for its specific fields.',C['text4'],'tiny').pack(side='left',padx=18,pady=4)

        tcard = ctk.CTkFrame(right, fg_color=C['card'], corner_radius=12, border_width=1, border_color=C['border'])
        tcard.pack(fill='both', expand=True, padx=14, pady=(2,14))
        tree_holder = {'frame':None, 'tree':None, 'cols':None}

        def current_columns():
            """Columns depend on which category is selected."""
            cname = filt['cname']
            if cname and cname in IT_CATEGORY_FIELDS:
                flds = IT_CATEGORY_FIELDS[cname]
                # cap at ~6 shown fields to keep table readable
                show = flds[:6]
                return ['ID'] + show + ['Stock']
            # All-assets generic view
            return ['ID','Name','Category','Model','Serial','Location','Stock']

        def rebuild_table():
            # destroy old tree, build a fresh one with the right columns
            if tree_holder['frame'] is not None:
                tree_holder['frame'].destroy()
            cols = current_columns()
            widths = []
            for c in cols:
                if c=='ID': widths.append(50)
                elif c=='Stock': widths.append(64)
                elif c in ('Name','Description','Specs','Specification'): widths.append(200)
                else: widths.append(130)
            cbt = CheckboxTree(tcard, tuple(cols), widths, height=16, simple_mode=True)
            cbt.frame.pack(fill='both', expand=True, padx=2, pady=2)
            tree_holder['frame']=cbt.frame; tree_holder['tree']=cbt; tree_holder['cols']=cols
            self._it_tree = cbt
            cbt.tree.bind('<Double-1>', lambda e: self._it_asset_dlg(reload, True))
            load()

        def load():
            cbt = tree_holder['tree']; cols = tree_holder['cols']
            if cbt is None: return
            cbt.clear()
            srch = f"%{sv.get()}%"
            try:
                if filt['c']:
                    rows = qy("""SELECT a.AssetID, a.Name, ISNULL(a.Model,''), ISNULL(a.Serial,''),
                                        ISNULL(a.Location,''), a.Stock, ISNULL(a.DetailsJSON,''), ISNULL(c.Name,'-')
                                 FROM ITAssets a LEFT JOIN ITCategories c ON a.ITCatID=c.ITCatID
                                 WHERE a.ITCatID=? AND a.Name LIKE ? ORDER BY a.Name""",(filt['c'],srch))
                else:
                    rows = qy("""SELECT a.AssetID, a.Name, ISNULL(a.Model,''), ISNULL(a.Serial,''),
                                        ISNULL(a.Location,''), a.Stock, ISNULL(a.DetailsJSON,''), ISNULL(c.Name,'-')
                                 FROM ITAssets a LEFT JOIN ITCategories c ON a.ITCatID=c.ITCatID
                                 WHERE a.Name LIKE ? ORDER BY c.Name,a.Name""",(srch,))
            except Exception as e:
                messagebox.showerror('Load error', f'Could not load IT assets:\n{e}', parent=self)
                return
            for i,r in enumerate(rows):
                aid,name,model,serial,location,stock,djson,cat = r
                details = {}
                if djson:
                    try: details=json.loads(djson)
                    except: details={}
                if filt['cname'] and filt['cname'] in IT_CATEGORY_FIELDS:
                    # category-specific columns: that category's fields + Stock (ID handled by CheckboxTree)
                    vals=[]
                    for f in IT_CATEGORY_FIELDS[filt['cname']][:6]:
                        vals.append(details.get(f,''))
                    vals.append(stock)
                else:
                    vals=[name,cat,model,serial,location,stock]
                cbt.insert(aid, tuple(vals))

        def reload():
            # full reload incl. sidebar counts
            self._reload(self._it_manage)

        sv.trace_add('write', lambda *_: load())
        rebuild_table()

    # ── IT add/edit asset dialog ──────────────────────────────
    def _it_asset_dlg(self, reload_fn, edit=False):
        """Category-FIRST asset dialog. Pick a category, then that category's
        exact fields (from IT_CATEGORY_FIELDS) appear to be filled in.
        Field values are stored as JSON in ITAssets.DetailsJSON; the first
        field is used as the asset's display Name."""
        import json
        C = T()
        editing = None
        existing = None
        if edit:
            editing = self._it_tree_selection()
            if editing is None: messagebox.showinfo('Select','Tick an asset (or click its row) to edit.',parent=self); return
            r = qy("SELECT Name,ITCatID,ISNULL(DetailsJSON,''),Stock,ISNULL(Unit,''),ISNULL(AssetTag,''),ISNULL(Location,'') FROM ITAssets WHERE AssetID=?",(editing,))
            if r: existing = r[0]

        w = ctk.CTkToplevel(self); w.title('Edit Asset' if editing else 'Add Asset')
        w.geometry('540x720'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,'Edit IT Asset' if editing else 'Add IT Asset',C['text'],'h2').pack(pady=(16,2))

        cats = qy("SELECT ITCatID,Name FROM ITCategories ORDER BY Name")
        opts = [r[1] for r in cats]; cmap = {r[1]:r[0] for r in cats}; rev = {r[0]:r[1] for r in cats}

        # ── Step 1: category picker (top, always visible) ──
        topfm = ctk.CTkFrame(w, fg_color=C['surface']); topfm.pack(fill='x', padx=28, pady=(6,0))
        caps_lbl(topfm, '① Select Category').pack(anchor='w', pady=(2,3))
        init_cat = rev.get(existing[1]) if (existing and existing[1] in rev) else (opts[0] if opts else '')
        cv = ctk.StringVar(value=init_cat)
        cat_box = ctk.CTkComboBox(topfm, values=opts, variable=cv, width=480, height=40,
                       fg_color=C['surface2'], border_color=C['border'],
                       button_color=C['gold'], button_hover_color=C['gold2'],
                       text_color=C['text'], font=F['body'], corner_radius=6,
                       dropdown_fg_color=C['surface2'], dropdown_text_color=C['text'])
        cat_box.pack()
        if editing:
            cat_box.configure(state='disabled')  # don't move an asset between schemas on edit

        L(topfm, '② Fill in the fields for this category', C['text4'], 'xs').pack(anchor='w', pady=(12,2))

        # ── Step 2: dynamic per-category fields (rebuilt on category change) ──
        fm = ctk.CTkScrollableFrame(w, fg_color=C['surface'], width=480, height=380)
        fm.pack(fill='both', expand=True, padx=28, pady=(0,4))

        field_widgets = {}
        existing_details = {}
        if existing and existing[2]:
            try: existing_details = json.loads(existing[2])
            except: existing_details = {}

        def fields_for(cat_name):
            return IT_CATEGORY_FIELDS.get(cat_name, IT_DEFAULT_FIELDS)

        def build_fields(*_):
            for ch in fm.winfo_children(): ch.destroy()
            field_widgets.clear()
            cat_name = cv.get()
            flds = fields_for(cat_name)
            for fl in flds:
                caps_lbl(fm, fl).pack(anchor='w', pady=(9,3))
                e = inp(fm, w=440)
                val = existing_details.get(fl, '')
                if val: e.insert(0, val)
                e.pack(); field_widgets[fl] = e
            # common Stock + Unit (every category tracks a quantity)
            caps_lbl(fm, 'Stock / Quantity').pack(anchor='w', pady=(9,3))
            se = inp(fm, w=440); se.insert(0, str(existing[3]) if existing else '1'); se.pack()
            field_widgets['__stock__'] = se
            caps_lbl(fm, 'Unit').pack(anchor='w', pady=(9,3))
            ue = inp(fm, w=440); ue.insert(0, (existing[4] if existing else 'pcs')); ue.pack()
            field_widgets['__unit__'] = ue

        cv.trace_add('write', build_fields)
        build_fields()

        def save():
            try:
                cat_name = cv.get()
                if not cat_name: raise ValueError('Pick a category first')
                cid = cmap.get(cat_name)
                flds = fields_for(cat_name)
                details = {}
                for fl in flds:
                    details[fl] = field_widgets[fl].get().strip()
                # display name = first field's value (fallback to category)
                name = details.get(flds[0], '').strip() or cat_name
                try: stock = int(float(field_widgets['__stock__'].get() or 0))
                except: stock = 0
                unit = field_widgets['__unit__'].get().strip() or 'pcs'
                # convenience columns for search/display
                asset_tag = details.get('Asset Tag') or details.get('Tag') or details.get('ASSET TAG') or ''
                model = details.get('Model') or details.get('Model No.') or ''
                serial = details.get('Serial') or details.get('S/N') or details.get('SN') or details.get('Serial Number') or details.get('Serial No.') or ''
                location = details.get('Location') or details.get('Default Location') or details.get('Class/Office') or ''
                djson = json.dumps(details, ensure_ascii=False)
                if editing:
                    ex("""UPDATE ITAssets SET Name=?,Brand=?,Model=?,Serial=?,Specification=?,ITCatID=?,Stock=?,Unit=?,
                          DetailsJSON=?,AssetTag=?,Location=? WHERE AssetID=?""",
                       (name,'',model,serial,'',cid,stock,unit,djson,asset_tag,location,editing))
                else:
                    ex("""INSERT INTO ITAssets(Name,Brand,Model,Serial,Specification,ITCatID,Stock,Unit,DetailsJSON,AssetTag,Location)
                          VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                       (name,'',model,serial,'',cid,stock,unit,djson,asset_tag,location))
                reload_fn(); w.destroy()
            except Exception as e:
                messagebox.showerror('Error',str(e),parent=w)
        blue_btn(w,'Save Asset',save,w=480,h=44).pack(pady=(4,16),padx=30)

    def _it_asset_dlg_for(self, cat_id, cat_name, reload_fn, edit=False):
        """Add/Edit dialog already scoped to ONE category (called from a
        category mode-screen). If cat_id is None (the 'All Assets' screen),
        falls back to the full picker dialog."""
        if cat_id is None:
            return self._it_asset_dlg(reload_fn, edit)
        import json
        C = T()
        editing = None
        existing = None
        if edit:
            editing = self._it_tree_selection()
            if editing is None: messagebox.showinfo('Select','Tick an asset (or click its row) to edit.',parent=self); return
            r = qy("SELECT Name,ITCatID,ISNULL(DetailsJSON,''),Stock,ISNULL(Unit,''),ISNULL(AssetTag,''),ISNULL(Location,'') FROM ITAssets WHERE AssetID=?",(editing,))
            if r: existing = r[0]

        w = ctk.CTkToplevel(self); w.title(f'Edit {cat_name} Asset' if editing else f'Add {cat_name} Asset')
        w.geometry('540x680'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,'Edit Asset' if editing else 'Add Asset',C['text'],'h2').pack(pady=(16,2))
        L(w, cat_name, C['gold'], 'sm').pack(pady=(0,8))

        fm = ctk.CTkScrollableFrame(w, fg_color=C['surface'], width=480, height=420)
        fm.pack(fill='both', expand=True, padx=28, pady=(0,4))

        existing_details = {}
        if existing and existing[2]:
            try: existing_details = json.loads(existing[2])
            except: existing_details = {}

        flds = IT_CATEGORY_FIELDS.get(cat_name, IT_DEFAULT_FIELDS)
        field_widgets = {}
        for fl in flds:
            caps_lbl(fm, fl).pack(anchor='w', pady=(9,3))
            e = inp(fm, w=440)
            val = existing_details.get(fl, '')
            if val: e.insert(0, val)
            e.pack(); field_widgets[fl] = e
        caps_lbl(fm, 'Stock / Quantity').pack(anchor='w', pady=(9,3))
        se = inp(fm, w=440); se.insert(0, str(existing[3]) if existing else '1'); se.pack()
        caps_lbl(fm, 'Unit').pack(anchor='w', pady=(9,3))
        ue = inp(fm, w=440); ue.insert(0, (existing[4] if existing else 'pcs')); ue.pack()

        def save():
            try:
                details = {fl: field_widgets[fl].get().strip() for fl in flds}
                name = details.get(flds[0], '').strip() or cat_name
                try: stock = int(float(se.get() or 0))
                except: stock = 0
                unit = ue.get().strip() or 'pcs'
                asset_tag = details.get('Asset Tag') or details.get('Tag') or ''
                model = details.get('Model') or details.get('Model No.') or ''
                serial = (details.get('Serial') or details.get('S/N') or details.get('SN')
                          or details.get('Serial Number') or details.get('Serial No.') or '')
                location = details.get('Location') or details.get('Default Location') or details.get('Class/Office') or ''
                djson = json.dumps(details, ensure_ascii=False)
                if editing:
                    ex("""UPDATE ITAssets SET Name=?,Model=?,Serial=?,ITCatID=?,Stock=?,Unit=?,
                          DetailsJSON=?,AssetTag=?,Location=? WHERE AssetID=?""",
                       (name,model,serial,cat_id,stock,unit,djson,asset_tag,location,editing))
                else:
                    ex("""INSERT INTO ITAssets(Name,Brand,Model,Serial,Specification,ITCatID,Stock,Unit,DetailsJSON,AssetTag,Location)
                          VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                       (name,'',model,serial,'',cat_id,stock,unit,djson,asset_tag,location))
                reload_fn(); w.destroy()
            except Exception as e:
                messagebox.showerror('Error',str(e),parent=w)
        blue_btn(w,'Save Asset',save,w=480,h=44).pack(pady=(4,16),padx=30)

    def _it_bulk_one(self, cat_name, reload_fn):
        """Bulk upload scoped to ONE category — pick any Excel file, but only
        the sheet matching this category's name gets imported."""
        import json
        path = filedialog.askopenfilename(title=f'Select Excel containing the "{cat_name}" sheet',
                filetypes=[('Excel','*.xlsx *.xls')], parent=self)
        if not path: return
        try:
            import openpyxl
        except:
            messagebox.showerror('Missing','Run: pip install openpyxl',parent=self); return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception as e:
            messagebox.showerror('Error',f'Could not read file: {e}',parent=self); return
        if cat_name not in wb.sheetnames:
            messagebox.showerror('Sheet not found', f'This file has no sheet named "{cat_name}".', parent=self); return

        added = self._it_import_sheet(wb, cat_name, reload_fn, show_message=False)
        if added is None:
            messagebox.showwarning('Could not import',
                f'The "{cat_name}" sheet doesn\'t have a clean single-table layout '
                f'(it may have multiple blocks or notes). Add these assets manually instead.', parent=self)
        else:
            messagebox.showinfo('Bulk Upload', f'Imported {added} asset(s) into {cat_name}.', parent=self)
            self.after(50, reload_fn)

    # ── IT delete ─────────────────────────────────────────────
    def _it_del(self, reload_fn):
        ids = self._it_tree_selection(multi=True)
        if not ids: messagebox.showinfo('Select','Tick the asset(s) you want to delete (or click a row).',parent=self); return
        C = T()
        w = ctk.CTkToplevel(self); w.title('Confirm Delete')
        w.geometry('380x200'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['red']).pack(fill='x')
        L(w,'Delete IT assets?',C['text'],'h2').pack(pady=(18,8))
        L(w,f'{len(ids)} asset(s) will be permanently deleted.',C['text3'],'sm').pack()
        L(w,'This only affects IT assets — not Procurement.',C['text4'],'tiny').pack(pady=(4,0))
        bf = ctk.CTkFrame(w,fg_color=C['surface']); bf.pack(pady=18)
        ghost_btn(bf,'Cancel',w.destroy,w=110,h=38).pack(side='left',padx=8)
        def confirm():
            for i in ids: ex("DELETE FROM ITAssets WHERE AssetID=?",(i,))
            reload_fn(); w.destroy()
        red_btn(bf,'Delete All',confirm,w=120,h=38).pack(side='left',padx=8)

    def _it_tree_selection(self, multi=False):
        """Return checked asset ID(s) from the IT assets CheckboxTree, falling
        back to the highlighted row if nothing is ticked. multi=True returns a
        list (for delete); multi=False returns a single ID or None (for edit)."""
        cbt = getattr(self, '_it_tree', None)
        if cbt is None: return [] if multi else None
        ids = list(cbt.checked)
        if not ids:
            sel = cbt.tree.selection()
            if sel:
                iid = cbt._id_map.get(sel[0])
                if iid is not None: ids = [iid]
        if multi: return ids
        return ids[0] if ids else None

    # ── IT bulk upload from Excel ─────────────────────────────
    def _it_find_header_row(self, rows, expected):
        """Return (row_index, match_count) for the row that best matches the
        expected field names — using EXACT or near-exact matches only."""
        exp = [e.lower() for e in expected]
        best_i, best_score = None, 0
        for i, row in enumerate(rows[:12]):
            cells = [str(c).strip().lower() if c is not None else '' for c in row]
            score = 0
            for c in cells:
                if not c or len(c) < 2: continue
                for e in exp:
                    if c == e:
                        score += 2; break
                    if (c in e or e in c) and abs(len(c)-len(e)) <= 4:
                        score += 1; break
            if score > best_score:
                best_score, best_i = score, i
        return best_i, best_score

    def _it_build_col_map(self, header_cells, fields):
        """Map column index -> field name using STRICT matching only.
        Exact matches win first; fuzzy matches are a fallback that prefers
        the closest-length field (so 'Model No.' doesn't get mistakenly
        claimed by the field 'Model'). Returns None on a duplicate field
        claim (a sign of a repeated/multi-block header)."""
        col_to_field = {}
        used_fields = {}
        remaining = []
        for ci, hc in enumerate(header_cells):
            hc = (hc or '').strip()
            if not hc: continue
            hcl = hc.lower()
            exact = next((f for f in fields if f.lower() == hcl), None)
            if exact:
                if exact in used_fields:
                    return None
                used_fields[exact] = ci
                col_to_field[ci] = exact
            else:
                remaining.append((ci, hc, hcl))
        for ci, hc, hcl in remaining:
            candidates = []
            for f in fields:
                if f in used_fields: continue
                fl = f.lower()
                if (hcl in fl or fl in hcl) and abs(len(hcl)-len(fl)) <= 4:
                    candidates.append(f)
            if not candidates: continue
            best = min(candidates, key=lambda f: abs(len(f)-len(hc)))
            used_fields[best] = ci
            col_to_field[ci] = best
        return col_to_field if used_fields else None

    JUNK_ROW_WORDS = {'total', 'totals', 'grand total'}

    def _it_import_sheet(self, wb, sheet, reload_fn=None, show_message=True):
        """Import ONE clean sheet into ITAssets. The workbook is normalized:
        header row is row 1, data follows, one clean table per sheet.
        Returns the number of assets added, or None if the sheet header
        doesn't match the expected category fields."""
        import json
        fields = IT_CATEGORY_FIELDS.get(sheet)
        if not fields:
            return None
        ws = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return None

        # Header is row 0 in the clean workbook. Map header text -> column index.
        header_cells = [str(c).strip() if c is not None else '' for c in rows[0]]
        # Build a tolerant column map: exact match first, then case-insensitive
        col_for = {}
        low_header = [h.lower() for h in header_cells]
        for f in fields:
            idx = None
            if f in header_cells:
                idx = header_cells.index(f)
            elif f.lower() in low_header:
                idx = low_header.index(f.lower())
            if idx is not None:
                col_for[f] = idx
        if not col_for:
            return None

        cat_map = {r[1].lower(): r[0] for r in qy("SELECT ITCatID,Name FROM ITCategories")}
        cid = cat_map.get(sheet.lower())
        if not cid:
            ex("INSERT INTO ITCategories(Name) VALUES(?)",(sheet,))
            cid = qy("SELECT TOP 1 ITCatID FROM ITCategories ORDER BY ITCatID DESC")[0][0]

        added = 0
        for row in rows[1:]:
            if not row or all(v is None for v in row): continue
            cells = [str(c).strip() if c is not None else '' for c in row]
            details = {}
            for f, ci in col_for.items():
                val = cells[ci] if ci < len(cells) else ''
                if val:
                    details[f] = val
            if not details:
                continue
            name = details.get(fields[0], '') or next(iter(details.values()), sheet)
            model = details.get('Model') or details.get('Model No.') or ''
            serial = (details.get('Serial') or details.get('S/N') or details.get('SN')
                      or details.get('Serial Number') or details.get('Serial No.') or '')
            asset_tag = details.get('Asset Tag') or details.get('Tag') or ''
            location = (details.get('Location') or details.get('Default Location')
                        or details.get('Class/Office') or '')
            ex("""INSERT INTO ITAssets(Name,Brand,Model,Serial,Specification,ITCatID,Stock,Unit,DetailsJSON,AssetTag,Location)
                  VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
               (name,'',model,serial,'',cid,1,'pcs',json.dumps(details,ensure_ascii=False),asset_tag,location))
            added += 1
        return added

    def _it_bulk(self, reload_fn):
        """Ask the user which import engine to use, then run it.
        TWO engines:
          • Clean engine — for the normalized 'BSB_IT_Clean.xlsx' (header row 1)
          • Original engine — for the raw 'BSB - IT Inventory' workbook with
            stacked tables, header offsets and side-by-side blocks.
        This lets you compare which gives better results on your file."""
        C = T()
        w = ctk.CTkToplevel(self); w.title('IT Bulk Upload')
        w.geometry('500x340'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,'IT Bulk Upload',C['text'],'h2').pack(pady=(18,4))
        L(w,'Pick which workbook format you are uploading.',C['text3'],'sm').pack(pady=(0,14))
        fm=ctk.CTkFrame(w,fg_color=C['surface']); fm.pack(fill='both',expand=True,padx=30)

        def pick(mode):
            w.destroy()
            if mode=='clean': self._it_bulk_clean(reload_fn)
            else: self._it_bulk_original(reload_fn)

        card1=ctk.CTkFrame(fm,fg_color=C['card'],corner_radius=10,border_width=1,border_color=C['border'])
        card1.pack(fill='x',pady=6)
        L(card1,'✅  Clean workbook  (recommended)',C['text'],'sm_b').pack(anchor='w',padx=14,pady=(12,2))
        L(card1,'Use the BSB_IT_Clean.xlsx I generated — one clean table per sheet.',C['text3'],'tiny',wraplength=400,justify='left').pack(anchor='w',padx=14,pady=(0,10))
        gold_btn(card1,'Use Clean Engine',lambda: pick('clean'),w=180,h=34).pack(anchor='e',padx=14,pady=(0,12))

        card2=ctk.CTkFrame(fm,fg_color=C['card'],corner_radius=10,border_width=1,border_color=C['border'])
        card2.pack(fill='x',pady=6)
        L(card2,'📄  Original workbook',C['text'],'sm_b').pack(anchor='w',padx=14,pady=(12,2))
        L(card2,'Use your raw BSB - IT Inventory file directly (stacked/messy tables).',C['text3'],'tiny',wraplength=400,justify='left').pack(anchor='w',padx=14,pady=(0,10))
        ghost_btn(card2,'Use Original Engine',lambda: pick('original'),w=180,h=34).pack(anchor='e',padx=14,pady=(0,12))

    def _it_bulk_clean(self, reload_fn):
        """Engine A: import from the normalized workbook (header row 1)."""
        path = filedialog.askopenfilename(title='Select the CLEAN BSB IT Inventory Excel',
                filetypes=[('Excel','*.xlsx *.xls')], parent=self)
        if not path: return
        try:
            import openpyxl
        except:
            messagebox.showerror('Missing','Run: pip install openpyxl',parent=self); return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception as e:
            messagebox.showerror('Error',f'Could not read file: {e}',parent=self); return

        total_added = 0; sheets_done = 0; skipped_sheets = []
        for sheet in wb.sheetnames:
            if sheet.strip().lower() in ('summary','xdo_metadata'): continue
            if sheet not in IT_CATEGORY_FIELDS:
                skipped_sheets.append(sheet); continue
            added = self._it_import_sheet(wb, sheet)
            if not added:
                skipped_sheets.append(sheet)
            else:
                total_added += added; sheets_done += 1

        msg = f'[Clean engine] Imported {total_added} asset(s) across {sheets_done} categor{"y" if sheets_done==1 else "ies"}.'
        if skipped_sheets:
            msg += (f'\n\nSkipped {len(skipped_sheets)} sheet(s):\n'
                    + ", ".join(skipped_sheets[:10]) + (' …' if len(skipped_sheets)>10 else ''))
        messagebox.showinfo('IT Bulk Upload', msg, parent=self)
        self.after(50, reload_fn)

    def _it_bulk_original(self, reload_fn):
        """Engine B: import directly from the RAW messy workbook.
        Maps original sheet names to clean category names, finds the real
        header row even when offset, cuts off stacked sub-tables, and handles
        the side-by-side (Printers, PCLaptop) and color-status (VR, Macbooks)
        special cases — all without needing the workbook to be pre-cleaned."""
        path = filedialog.askopenfilename(title='Select the ORIGINAL BSB IT Inventory Excel',
                filetypes=[('Excel','*.xlsx *.xls')], parent=self)
        if not path: return
        try:
            import openpyxl
        except:
            messagebox.showerror('Missing','Run: pip install openpyxl',parent=self); return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            wb_s = openpyxl.load_workbook(path, data_only=False)  # for highlight colors
        except Exception as e:
            messagebox.showerror('Error',f'Could not read file: {e}',parent=self); return

        total_added = 0; per = {}; skipped = []
        for raw_sheet, fn in self._ORIG_SHEET_PARSERS.items():
            if raw_sheet not in wb.sheetnames:
                continue
            try:
                added = fn(self, wb, wb_s, raw_sheet)
            except Exception as e:
                skipped.append(f'{raw_sheet} ({e})'); continue
            if added:
                total_added += added; per[raw_sheet] = added
            else:
                skipped.append(raw_sheet)

        lines = [f'[Original engine] Imported {total_added} asset(s).', '']
        for k,v in per.items(): lines.append(f'  • {k}: {v}')
        if skipped:
            lines += ['', 'Skipped: ' + ', '.join(s for s in skipped[:12])]
        messagebox.showinfo('IT Bulk Upload', '\n'.join(lines), parent=self)
        self.after(50, reload_fn)

    # ─── helpers for the ORIGINAL-workbook engine ───
    def _orig_cat_id(self, clean_name):
        """Resolve (or create) the ITCategories row for a clean category name."""
        cat_map = {r[1].lower(): r[0] for r in qy("SELECT ITCatID,Name FROM ITCategories")}
        cid = cat_map.get(clean_name.lower())
        if not cid:
            ex("INSERT INTO ITCategories(Name) VALUES(?)",(clean_name,))
            cid = qy("SELECT TOP 1 ITCatID FROM ITCategories ORDER BY ITCatID DESC")[0][0]
        return cid

    def _orig_insert(self, cid, details, cat_name):
        import json
        flds = IT_CATEGORY_FIELDS.get(cat_name, IT_DEFAULT_FIELDS)
        name = details.get(flds[0],'') or next(iter(details.values()), cat_name)
        model = details.get('Model') or details.get('Model No.') or details.get('Device Type') or ''
        serial = (details.get('Serial') or details.get('S/N') or details.get('SN')
                  or details.get('Serial Number') or details.get('Serial No.') or '')
        asset_tag = details.get('Asset Tag') or details.get('Tag') or ''
        location = (details.get('Location') or details.get('Default Location')
                    or details.get('Class/Office') or '')
        ex("""INSERT INTO ITAssets(Name,Brand,Model,Serial,Specification,ITCatID,Stock,Unit,DetailsJSON,AssetTag,Location)
              VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
           (name,'',model,serial,'',cid,1,'pcs',json.dumps(details,ensure_ascii=False),asset_tag,location))

    @staticmethod
    def _fill_color(wb_s, sheet, row, col=2):
        try:
            c=wb_s[sheet].cell(row,col); f=c.fill
            if f and f.patternType=='solid' and f.fgColor and f.fgColor.rgb not in ('00000000',None,'FFFFFFFF'):
                return f.fgColor.rgb
        except: pass
        return None

    def _orig_simple(self, wb, sheet, clean_name, header_idx, col_start, fields):
        """Generic parser: header at row `header_idx` (0-based), data follows,
        columns starting at `col_start` (0-based) map 1:1 to `fields`.
        Stops at first fully-blank row."""
        ws=wb[sheet]; rows=list(ws.iter_rows(values_only=True))
        cid=self._orig_cat_id(clean_name); added=0; blanks=0
        for i in range(header_idx+1, len(rows)):
            cells=[(str(c).strip() if c is not None else '') for c in rows[i]]
            window=cells[col_start:col_start+len(fields)]
            if not any(window):
                blanks+=1
                if blanks>=2: break
                continue
            blanks=0
            # skip a repeated header / section title
            if window and window[0] in ('CN','Sl No','Sl','ID','#') and not window[0].isdigit():
                continue
            details={fields[j]: window[j] for j in range(len(fields)) if j<len(window) and window[j]}
            if not details: continue
            # require at least the first real field present
            if not details.get(fields[0]) and len(details)<2: continue
            self._orig_insert(cid, details, clean_name); added+=1
        return added

    # ── per-sheet parsers for the ORIGINAL (raw) workbook ──
    # Each takes (self, wb, wb_s, sheet) and returns the count added.
    def _op_servers(self, wb, wb_s, sheet, clean):
        return self._orig_simple(wb, sheet, clean, 3, 1,
            ['Asset Name','Description','Roles','IP','Type of Server','Model','Specs','S/N'])
    def _op_active_servers(self, wb, wb_s, sheet): return self._op_servers(wb,wb_s,sheet,'Active Servers')
    def _op_decomm_servers(self, wb, wb_s, sheet): return self._op_servers(wb,wb_s,sheet,'Decomm Servers')

    def _op_network(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Network Devices', 4, 1,
            ['Device Name','Model','Management IP','Location','Access Methods','IOS Version'])

    def _op_ipphones(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'IP Phones', 2, 1,
            ['Ext. No','Name','MAC','SN','Asset Tag','Model','Brand','Remarks'])

    def _op_aps(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'APs', 4, 1,
            ['Asset Tag','Model No.','Manufacturer','Serial','MAC Address'])

    def _op_ups(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'UPS', 6, 1,
            ['Asset Name','Description','Roles','Model','S/N'])

    def _op_ipads(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'iPads', 3, 1,
            ['Asset Tag','Model','Model No.','Serial','Default Location','School','Floor','Building','Status'])

    def _op_mobile(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Mobile Devices', 3, 3,
            ['Department','Device Model','Asset Tag','Mobile Number','Package'])

    def _op_chromebooks(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Chromebooks', 4, 1,
            ['Asset Tag','Model','Model No.','Category','Manufacturer','Serial','Purchased','Supplier','Location','Warranty','School','Floor','Building'])

    def _op_desktops(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Desktops', 3, 1,
            ['Asset Tag','Model','Serial Number','Class/Office','Floor','Building'])

    def _op_laptops(self, wb, wb_s, sheet):
        n = self._orig_simple(wb, sheet, 'Laptops', 4, 1,
            ['Asset Tag','Model','Model No.','Category','Manufacturer','Serial','Default Location','Checked Out','Location'])
        return n

    def _op_projectors(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Projectors', 3, 1,
            ['Brand','Model','Class/Office','Floor','Building'])

    def _op_cameras(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Cameras', 3, 1,
            ['Asset Tag','Brand','Model','Class/Office','Floor','Building'])

    def _op_smartboards(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Smartboards', 3, 1,
            ['Asset Tag','Brand','S/N','Model','Class/Office','Activation Key','Building'])

    def _op_tv(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'TV', 3, 1,
            ['Brand','S/N','Model','Class/Office'])

    def _op_av(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Audio Visual devices', 3, 1,
            ['Model','Brand','Description','Location','Qty','Remarks'])

    def _op_kiosk(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Kiosk Machines', 3, 1,
            ['Asset Tag','Model','Serial Number','Device Status'])

    def _op_monitors(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Monitors n Docking Station', 6, 1,
            ['Name','Position','School/Department','Assigned Device','Serial Number','Model'])

    # ── special-case parsers ──
    def _op_cctv(self, wb, wb_s, sheet):
        """CCTV: import the real DEVICE table (ID/Device Type/IPv4/Software/Gateway/Serial)."""
        import json
        ws=wb[sheet]; rows=list(ws.iter_rows(values_only=True))
        fields=['ID','Device Type','IPv4 Address','Software Version','IPv4 Gateway','Serial Number']
        dev_start=None
        for i,row in enumerate(rows):
            cells=[(str(c).strip() if c is not None else '') for c in row[:6]]
            if cells[0]=='ID' and 'Device Type' in cells[1]:
                dev_start=i; break
        if dev_start is None: return 0
        cid=self._orig_cat_id('CCTV'); added=0
        for i in range(dev_start+1,len(rows)):
            cells=[(str(c).strip() if c is not None else '') for c in rows[i][:6]]
            if not any(cells): break
            if cells[0] and cells[0].replace('.','').isdigit() and cells[1]:
                details={fields[j]:cells[j] for j in range(len(fields)) if j<len(cells) and cells[j]}
                self._orig_insert(cid, details, 'CCTV'); added+=1
        return added

    def _op_infants_camera(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Infants Camera', 3, 1,
            ['Brand','Model','Serial Number'])

    def _op_access_control(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Access Control', 3, 2, ['Location','Count'])

    def _op_printers(self, wb, wb_s, sheet):
        """Printers: side-by-side tables. Left = Sl/Name/Model/Location/IP,
        Right (MFP) = CN/DeviceName/Location/Model/IP/Tag."""
        ws=wb[sheet]; rows=list(ws.iter_rows(values_only=True))
        cid=self._orig_cat_id('Printers'); added=0
        for i in range(7,len(rows)):
            c=[(str(x).strip() if x is not None else '') for x in rows[i][:13]]
            if c[1] and c[1] not in ('Printer Name','Total'):
                d={'Printer Name':c[1],'Model':c[2],'Location':c[3],'IP':c[4]}
                d={k:v for k,v in d.items() if v}
                if d: self._orig_insert(cid,d,'Printers'); added+=1
            elif c[2] and not c[1] and c[2] not in ('Model','Total'):
                d={'Model':c[2],'Location':c[3],'IP':c[4]}; d={k:v for k,v in d.items() if v}
                if d: self._orig_insert(cid,d,'Printers'); added+=1
            if len(c)>=12 and c[7] and c[7] not in ('Device Name',''):
                d={'Printer Name':c[7],'Model':c[9],'Location':c[8],'IP':c[10],'Asset Tag':c[11]}
                d={k:v for k,v in d.items() if v}
                if d: self._orig_insert(cid,d,'Printers'); added+=1
        return added

    def _op_vr(self, wb, wb_s, sheet):
        """VR: header at row 5; status from cell highlight (yellow=partial, red=broken)."""
        ws=wb[sheet]; rows=list(ws.iter_rows(values_only=True))
        fields=['Asset Tag','Model','Serial No.','Location','Bundle SN','Intune','Remarks']
        cid=self._orig_cat_id('VR'); added=0
        for i in range(6,len(rows)):
            c=[(str(x).strip() if x is not None else '') for x in rows[i][:8]]
            if not c[1]: continue
            col=self._fill_color(wb_s, sheet, i+1, 2)
            status='Broken' if col=='FFFF0000' else ('Partially Broken' if col=='FFFFFF00' else 'Working')
            details={fields[j]:c[j+1] for j in range(len(fields)) if j+1<len(c) and c[j+1]}
            details['Status']=status
            self._orig_insert(cid, details, 'VR'); added+=1
        return added

    def _op_macbooks(self, wb, wb_s, sheet):
        """Macbooks: header at row 4; red highlight or 'Decommission' note = Decommissioned."""
        ws=wb[sheet]; rows=list(ws.iter_rows(values_only=True))
        fields=['Asset Tag','Model No.','Manufacturer','Serial','Location','School','Floor','Building','RAM','CPU','HDD']
        cid=self._orig_cat_id('Macbooks & iMacs'); added=0
        for i in range(5,len(rows)):
            c=[(str(x).strip() if x is not None else '') for x in rows[i][:13]]
            if not c[1] or c[1]=='Asset Tag': continue
            col=self._fill_color(wb_s, sheet, i+1, 2)
            note=c[12] if len(c)>12 else ''
            status='Decommissioned' if (col=='FFFF0000' or 'decom' in note.lower()) else 'Active'
            details={fields[j]:c[j+1] for j in range(len(fields)) if j+1<len(c) and c[j+1]}
            details['Status']=status
            self._orig_insert(cid, details, 'Macbooks & iMacs'); added+=1
        return added

    def _op_new_teacher_laptops(self, wb, wb_s, sheet):
        """New Laptop for teachers → merge into Laptops."""
        ws=wb[sheet]; rows=list(ws.iter_rows(values_only=True))
        cid=self._orig_cat_id('Laptops'); added=0
        for i in range(10,len(rows)):
            c=[(str(x).strip() if x is not None else '') for x in rows[i][:8]]
            if c[4] and c[4] not in ('Assigned Device',''):
                d={'Asset Tag':c[4],'Model':c[6],'Category':'Teacher Laptop',
                   'Serial':c[5],'Default Location':c[3],'Checked Out':c[1],'Location':c[7]}
                d={k:v for k,v in d.items() if v}
                if d: self._orig_insert(cid,d,'Laptops'); added+=1
        return added

    def _op_pclaptop_decom(self, wb, wb_s, sheet):
        """PCLaptop to be decom: side-by-side Desktops + Laptops."""
        ws=wb[sheet]; rows=list(ws.iter_rows(values_only=True))
        cid=self._orig_cat_id('PC Laptop to Decom'); added=0
        for i in range(1,len(rows)):
            c=[(str(x).strip() if x is not None else '') for x in rows[i][:14]]
            if c[2] and c[2].startswith('BSBDT'):
                d={'Type':'Desktop','Tag':c[2],'Make/Model':c[3],'Model':c[3],'Serial':c[4]}
                self._orig_insert(cid,{k:v for k,v in d.items() if v},'PC Laptop to Decom'); added+=1
            if len(c)>=14 and c[9] and c[9].startswith('BSBNB'):
                d={'Type':'Laptop','Tag':c[9],'Make/Model':c[10],'Model':c[11],'Serial':c[12],'Location':c[13]}
                self._orig_insert(cid,{k:v for k,v in d.items() if v},'PC Laptop to Decom'); added+=1
        return added

    def _op_ipadmac_decom(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'iPads Macbooks to Decom', 3, 1,
            ['Asset Tag','Model','Serial Number','Device Status'])

    def _op_smartboard_decom(self, wb, wb_s, sheet):
        return self._orig_simple(wb, sheet, 'Smartboard to Decom', 1, 2,
            ['Model','Brand','Serial Number','Location'])

    # Maps RAW (original) sheet name → parser method.
    _ORIG_SHEET_PARSERS = {
        'Active Servers':            lambda s,wb,ws,sh: s._op_active_servers(wb,ws,sh),
        'Decomm Servers':            lambda s,wb,ws,sh: s._op_decomm_servers(wb,ws,sh),
        'Network Devices':           lambda s,wb,ws,sh: s._op_network(wb,ws,sh),
        'IP Phones':                 lambda s,wb,ws,sh: s._op_ipphones(wb,ws,sh),
        'CCTV':                      lambda s,wb,ws,sh: s._op_cctv(wb,ws,sh),
        'APs':                       lambda s,wb,ws,sh: s._op_aps(wb,ws,sh),
        'Printers':                  lambda s,wb,ws,sh: s._op_printers(wb,ws,sh),
        'UPS':                       lambda s,wb,ws,sh: s._op_ups(wb,ws,sh),
        'iPads':                     lambda s,wb,ws,sh: s._op_ipads(wb,ws,sh),
        'VR':                        lambda s,wb,ws,sh: s._op_vr(wb,ws,sh),
        'Mobile Devices':            lambda s,wb,ws,sh: s._op_mobile(wb,ws,sh),
        'Chromebooks':               lambda s,wb,ws,sh: s._op_chromebooks(wb,ws,sh),
        'Macbooks & iMacs':          lambda s,wb,ws,sh: s._op_macbooks(wb,ws,sh),
        'Desktops':                  lambda s,wb,ws,sh: s._op_desktops(wb,ws,sh),
        'Laptops':                   lambda s,wb,ws,sh: s._op_laptops(wb,ws,sh),
        'New Laptop for teachers':   lambda s,wb,ws,sh: s._op_new_teacher_laptops(wb,ws,sh),
        'Projectors':                lambda s,wb,ws,sh: s._op_projectors(wb,ws,sh),
        'Cameras':                   lambda s,wb,ws,sh: s._op_cameras(wb,ws,sh),
        'Infants camera':            lambda s,wb,ws,sh: s._op_infants_camera(wb,ws,sh),
        'Smartboards':               lambda s,wb,ws,sh: s._op_smartboards(wb,ws,sh),
        'TV':                        lambda s,wb,ws,sh: s._op_tv(wb,ws,sh),
        'Audio Visual devices':      lambda s,wb,ws,sh: s._op_av(wb,ws,sh),
        'Access Control':            lambda s,wb,ws,sh: s._op_access_control(wb,ws,sh),
        'Kiosk Machines':            lambda s,wb,ws,sh: s._op_kiosk(wb,ws,sh),
        'Monitors n Docking Station- New': lambda s,wb,ws,sh: s._op_monitors(wb,ws,sh),
        'PCLaptop to be decom':      lambda s,wb,ws,sh: s._op_pclaptop_decom(wb,ws,sh),
        'iPadsMacbooks to be decom': lambda s,wb,ws,sh: s._op_ipadmac_decom(wb,ws,sh),
        'Smarboard to be decom':     lambda s,wb,ws,sh: s._op_smartboard_decom(wb,ws,sh),
    }

    # ── IT categories tab ─────────────────────────────────────
    def _it_tab_cats(self, parent):
        C = T()
        main = ctk.CTkFrame(parent,fg_color=C['bg']); main.pack(fill='both',expand=True,padx=16,pady=14)
        p = ctk.CTkFrame(main,fg_color=C['card'],corner_radius=12,border_width=1,border_color=C['border'])
        p.pack(side='left',fill='both',expand=True,padx=7)
        L(p,'IT Asset Categories',C['text'],'h3').pack(pady=14)
        hline(p)
        box = ctk.CTkFrame(p,fg_color=C['card']); box.pack(fill='both',expand=True,padx=10,pady=8)
        cbt = CheckboxTree(box, ('Category','Assets'), [240,90], height=16, simple_mode=True)
        cbt.frame.pack(fill='both', expand=True)
        hline(p)
        bf = ctk.CTkFrame(p,fg_color=C['card']); bf.pack(pady=10)
        def lc():
            cbt.clear()
            for r in qy("SELECT ITCatID,Name FROM ITCategories ORDER BY Name"):
                cnt = (qy("SELECT COUNT(*) FROM ITAssets WHERE ITCatID=?",(r[0],)) or [(0,)])[0][0]
                cbt.insert(r[0], (r[1], cnt))
        def ac():
            n = self._ask('Add IT Category','Category name:')
            if n: ex("INSERT INTO ITCategories(Name) VALUES(?)",(n,)); lc()
        def dc():
            ids = list(cbt.checked)
            if not ids:
                sel = cbt.tree.selection()
                if sel:
                    iid = cbt._id_map.get(sel[0])
                    if iid is not None: ids=[iid]
            if not ids:
                messagebox.showinfo('Select','Tick the categories you want to delete (or click a row).',parent=self); return
            n=len(ids)
            if messagebox.askyesno('Delete',f'Delete {n} categor{"y" if n==1 else "ies"}? Assets keep but lose their category.',parent=self):
                for cid in ids: ex("DELETE FROM ITCategories WHERE ITCatID=?",(cid,))
                lc()
        blue_btn(bf,'+ Add',ac,w=110,h=36).pack(side='left',padx=4)
        red_btn(bf,'🗑 Del Selected',dc,w=140,h=36).pack(side='left',padx=4)
        lc()

    def _tab_stock(self, parent):
        C = T()
        wrap = ctk.CTkFrame(parent, fg_color=C['bg']); wrap.pack(fill='both', expand=True)
        filt = {'p':None,'s':None}
        def on_filter(p,s): filt['p']=p; filt['s']=s; load()
        self._sidebar(wrap, on_filter)

        right = ctk.CTkFrame(wrap, fg_color=C['bg']); right.pack(side='left', fill='both', expand=True)

        # Toolbar
        tb = ctk.CTkFrame(right, fg_color=C['topbar'], corner_radius=0, height=60)
        tb.pack(fill='x'); tb.pack_propagate(False)
        sv = ctk.StringVar()
        L(tb,'Search',C['text4'],'xs').pack(side='left',padx=(18,8),pady=16)
        ctk.CTkEntry(tb, textvariable=sv, placeholder_text='Search items...',
                    width=280, height=36, fg_color=C['surface2'], border_color=C['border'],
                    text_color=C['text'], placeholder_text_color=C['text4'],
                    font=F['body'], corner_radius=6).pack(side='left',pady=14)
        blue_btn(tb,'＋ Add Item',lambda: self._item_dlg(load),w=120,h=36).pack(side='right',padx=4,pady=12)
        gold_btn(tb,'⬆ Bulk Upload',lambda: self._proc_bulk(load),w=130,h=36).pack(side='right',padx=4,pady=12)
        ghost_btn(tb,'✎ Edit',lambda: self._item_dlg(load,True),w=90,h=36).pack(side='right',padx=4,pady=12)
        gold_btn(tb,'# Set Qty',lambda: self._qty_dlg(load),w=110,h=36).pack(side='right',padx=4,pady=12)
        red_btn(tb,'🗑 Delete',lambda: self._del_items(load),w=100,h=36).pack(side='right',padx=4,pady=12)

        # Bulk action bar (shows when items selected)
        bulk_bar = ctk.CTkFrame(right, fg_color=C['check_bg'], corner_radius=0, height=46)
        bulk_bar.pack_propagate(False)
        bulk_lbl = L(bulk_bar, '0 items selected', C['text'], 'sm_b')
        bulk_lbl.pack(side='left', padx=16, pady=10)
        
        def export_selected():
            sel = self._cbt.get_selected_values() if hasattr(self,'_cbt') else []
            if not sel: messagebox.showinfo('Nothing selected','Select items first.', parent=self); return
            self._export_selected_dlg(sel)
        
        gold_btn(bulk_bar,'📊 Export Selected',export_selected,w=160,h=34).pack(side='left',padx=8,pady=6)
        gold_btn(bulk_bar,'# Set Qty',lambda: self._qty_dlg(load),w=110,h=34).pack(side='left',padx=4,pady=6)
        red_btn(bulk_bar,'🗑 Delete Selected',lambda: self._del_items(load),w=140,h=34).pack(side='left',padx=4,pady=6)

        # Hint
        hint = ctk.CTkFrame(right, fg_color=C['bg'], height=26); hint.pack(fill='x'); hint.pack_propagate(False)
        L(hint,'ⓘ  Check items to select · Click "🖼 View" to see an item\'s image · Double-click a row to set quantity',C['text4'],'tiny').pack(side='left',padx=18,pady=4)

        # Checkbox table
        tcard = ctk.CTkFrame(right, fg_color=C['card'], corner_radius=12, border_width=1, border_color=C['border'])
        tcard.pack(fill='both', expand=True, padx=14, pady=(2,14))
        
        self._cbt = CheckboxTree(tcard,
            ('ID','Item Name','Category','Stock','Unit','Price BHD','Min','LPO No.','Invoice No.'),
            [50,230,140,110,80,90,50,110,110], height=16, image_col=True)
        self._cbt.frame.pack(fill='both', expand=True, padx=2, pady=2)
        self._cbt._dbl_callback = lambda e: self._qty_dlg(load)
        self._cbt.on_image_click = lambda item_id: self._show_item_image(item_id)

        def update_bulk():
            n = len(self._cbt.checked)
            if n > 0:
                bulk_bar.pack(fill='x', after=tb)
                bulk_lbl.configure(text=f'{n} item{"s" if n>1 else ""} selected')
            else:
                bulk_bar.pack_forget()
            # Schedule next check
            right.after(200, update_bulk)
        right.after(200, update_bulk)

        def load():
            self._cbt.clear()
            base = """SELECT i.ItemID, i.Name, ISNULL(s.Name,'-'),
                             i.Quantity, i.Unit, i.UnitPrice, i.MinStock,
                             ISNULL(i.LPONumber,''), ISNULL(i.InvoiceNumber,'')
                      FROM Items i LEFT JOIN SubCategories s ON i.SubCatID=s.SubCatID """
            srch = f"%{sv.get()}%"
            if filt['s']:   rows = qy(base+"WHERE i.SubCatID=? AND i.Name LIKE ? ORDER BY i.Name",(filt['s'],srch))
            elif filt['p']: rows = qy(base+"WHERE s.ParentCatID=? AND i.Name LIKE ? ORDER BY s.Name,i.Name",(filt['p'],srch))
            else:           rows = qy(base+"WHERE i.Name LIKE ? ORDER BY s.Name,i.Name",(srch,))
            for r in rows:
                st, _ = status_info(r[3], r[6])
                self._cbt.insert(r[0], r, st)

        sv.trace_add('write', lambda *_: load())
        load()

    def _proc_bulk(self, reload_fn):
        """Bulk upload stationery items into Procurement's Items table from Excel.

        Supports TWO layouts (auto-detected):
        A) CATEGORY-HEADER layout (your stationary_list.xlsx):
              A row with only column A filled = a CATEGORY heading.
              Item rows below it: Name | Unit | Price ('BHD 0.695')  belong to that category.
        B) COLUMN layout (template):
              Item Name | Category | Quantity | Unit | Price | Min Stock  (with a header row)
        """
        path = filedialog.askopenfilename(title='Select Stationery Excel',
                filetypes=[('Excel','*.xlsx *.xls')], parent=self)
        if not path: return
        try:
            import openpyxl
        except:
            messagebox.showerror('Missing','Run: pip install openpyxl',parent=self); return
        try:
            wb = openpyxl.load_workbook(path, data_only=True); ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
        except Exception as e:
            messagebox.showerror('Error',f'Could not read file: {e}',parent=self); return

        def parse_price(v):
            if v is None: return 0.0
            s = str(v).upper().replace('BHD','').replace('BD','').replace(',','').strip()
            try: return float(s)
            except: return 0.0
        def to_int(v, default=0):
            try: return int(float(str(v).strip()))
            except: return default

        # Detect layout: does the first non-empty row look like a column header
        # with a "category" column? (layout B) Otherwise use category-header (layout A).
        header_cells = []
        for r in rows:
            if r and any(v is not None and str(v).strip() for v in r):
                header_cells = [str(c).strip().lower() if c is not None else '' for c in r]
                break
        layout_b = ('category' in header_cells)

        sub_map = {r[1].lower(): r[0] for r in qy("SELECT SubCatID,Name FROM SubCategories")}
        par = qy("SELECT TOP 1 ParentCatID FROM ParentCategories ORDER BY ParentCatID")
        default_parent = par[0][0] if par else None
        # If no parent category exists at all, create one so subcategories have a home
        if default_parent is None:
            ex("INSERT INTO ParentCategories(Name) VALUES(?)",('STATIONERY',))
            default_parent = qy("SELECT TOP 1 ParentCatID FROM ParentCategories ORDER BY ParentCatID DESC")[0][0]

        def get_sub(cat_name):
            """Return SubCatID for a category name, creating it if needed."""
            if not cat_name: return None
            key = cat_name.lower()
            sid = sub_map.get(key)
            if not sid:
                ex("INSERT INTO SubCategories(Name,ParentCatID) VALUES(?,?)",(cat_name, default_parent))
                sid = qy("SELECT TOP 1 SubCatID FROM SubCategories ORDER BY SubCatID DESC")[0][0]
                sub_map[key] = sid
            return sid

        added = 0; skipped = 0; cats_made = set()
        JUNK = {'total', 'grand total', ''}

        if layout_b:
            # ── Column layout ──
            for r in rows:
                if not r or all(v is None for v in r): continue
                cells = [str(c).strip() if c is not None else '' for c in r]
                if cells[0].lower() in ('item name','name','itemname'): continue
                name = cells[0]
                if not name: skipped += 1; continue
                cat_name = cells[1] if len(cells)>1 else ''
                qty  = to_int(cells[2]) if len(cells)>2 else 0
                unit = cells[3] if len(cells)>3 else 'pcs'
                price= parse_price(cells[4]) if len(cells)>4 else 0.0
                mn   = to_int(cells[5], 5) if len(cells)>5 else 5
                sid = get_sub(cat_name)
                if cat_name: cats_made.add(cat_name)
                ex("INSERT INTO Items(Name,SubCatID,Quantity,Unit,UnitPrice,MinStock) VALUES(?,?,?,?,?,?)",
                   (name, sid, qty, unit, price, mn))
                added += 1
        else:
            # ── Category-header layout (your real sheet) ──
            current_cat = None
            for r in rows:
                if not r or all(v is None for v in r): continue
                a = str(r[0]).strip() if r[0] is not None else ''
                b = str(r[1]).strip() if len(r)>1 and r[1] is not None else ''
                c = str(r[2]).strip() if len(r)>2 and r[2] is not None else ''
                if not a: continue
                # A heading row = column A filled, B and C both empty
                if not b and not c:
                    if a.lower() in JUNK:   # ignore "Total" footer rows
                        current_cat = None
                        continue
                    current_cat = a
                    get_sub(current_cat)    # create the category now
                    cats_made.add(current_cat)
                    continue
                # Otherwise it's an item row: Name | Unit | Price
                name = a
                unit = b or 'pcs'
                price = parse_price(c)
                sid = get_sub(current_cat) if current_cat else None
                ex("INSERT INTO Items(Name,SubCatID,Quantity,Unit,UnitPrice,MinStock) VALUES(?,?,?,?,?,?)",
                   (name, sid, 0, unit, price, 5))
                added += 1

        msg = f'Imported {added} item(s) across {len(cats_made)} categor{"y" if len(cats_made)==1 else "ies"}.'
        if skipped: msg += f'\nSkipped {skipped} blank row(s).'
        msg += '\n\nAll items start at quantity 0 — set real counts with "Set Qty".'
        messagebox.showinfo('Bulk Upload', msg, parent=self)
        reload_fn()

    def _export_selected_dlg(self, items):
        """Dialog asking which report to export for selected items"""
        C = T()
        w = ctk.CTkToplevel(self); w.title('Export Selected Items')
        w.geometry('440x380'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w, height=3, fg_color=C['gold']).pack(fill='x')
        L(w, f'Export {len(items)} Selected Item(s)', C['text'], 'h2').pack(pady=(18,4))
        L(w, 'Choose what to export:', C['text3'], 'sm').pack(pady=(0,20))

        fm = ctk.CTkFrame(w, fg_color=C['surface']); fm.pack(fill='x', padx=32)
        
        def do_export(report_type):
            folder = filedialog.askdirectory(title='Select save folder', parent=w)
            if not folder: return
            self._do_selected_export(items, report_type, folder)
            w.destroy()
        
        for icon, label, desc, rtype in [
            ('📦','Stock Report','Current quantity, unit & price for selected items','stock'),
            ('⚠️','Low Stock Check','Only show which selected items are low/out','low'),
            ('📋','Full Details','All details including min stock and category','full'),
        ]:
            card = ctk.CTkFrame(fm, fg_color=C['surface2'], corner_radius=8,
                               border_width=1, border_color=C['border'])
            card.pack(fill='x', pady=5)
            inner = ctk.CTkFrame(card, fg_color='transparent'); inner.pack(fill='x', padx=14, pady=10)
            L(inner, f'{icon}  {label}', C['text'], 'sm_b').pack(side='left')
            L(inner, desc, C['text3'], 'tiny').pack(side='left', padx=8)
            blue_btn(inner, 'Export', lambda rt=rtype: do_export(rt), w=80, h=30).pack(side='right')

    def _do_selected_export(self, items, report_type, folder):
        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        except:
            messagebox.showerror('Missing', 'Run: pip install openpyxl', parent=self); return
        
        C = T()
        wb = openpyxl.Workbook(); ws = wb.active
        thin = Side(style='thin', color='2563EB')
        bd = Border(left=thin,right=thin,top=thin,bottom=thin)
        ctr = Alignment(horizontal='center', vertical='center')
        navy = PatternFill('solid', fgColor='0B1F3A')
        gold = PatternFill('solid', fgColor='C9972C')
        
        if report_type == 'stock':
            ws.title = 'Selected Stock'
            headers = ['#','Item Name','Category','Quantity','Unit','Price BHD']
            title = f'BSB SELECTED ITEMS STOCK REPORT ({len(items)} items)'
        elif report_type == 'low':
            ws.title = 'Low Stock Check'
            headers = ['#','Item Name','Category','Quantity','Min Stock','Status']
            title = f'BSB SELECTED ITEMS - LOW STOCK CHECK ({len(items)} items)'
        else:
            ws.title = 'Full Details'
            headers = ['#','Item Name','Category','Quantity','Unit','Price BHD','Min Stock']
            title = f'BSB SELECTED ITEMS FULL DETAILS ({len(items)} items)'

        n = len(headers)
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n)
        c = ws.cell(1,1,title); c.fill=gold; c.font=Font('Calibri',bold=True,color='0B1F3A',size=13); c.alignment=ctr
        ws.cell(2,1,f'Generated: {datetime.now().strftime("%d/%m/%Y %H:%M")} · {self.current_user}').font=Font(italic=True,color='888888',size=10)
        for ci,h in enumerate(headers,1):
            c=ws.cell(3,ci,h); c.fill=navy; c.font=Font('Calibri',bold=True,color='FFFFFF',size=11); c.alignment=ctr; c.border=bd

        for ri, r in enumerate(items, 4):
            iid, name, cat, qty, unit, price, mn = r[:7]
            st, slabel = status_info(qty, mn)
            fill_color = 'FFEDED' if st=='out' else 'FFF4E5' if st=='low' else ('EAF1FE' if ri%2==0 else 'FFFFFF')
            fill = PatternFill('solid', fgColor=fill_color)
            bf = Font('Calibri', size=10, color='222222')
            
            if report_type == 'stock':
                row_data = [ri-3, name, cat, qty, unit, float(price)]
            elif report_type == 'low':
                row_data = [ri-3, name, cat, qty, mn, slabel]
            else:
                row_data = [ri-3, name, cat, qty, unit, float(price), mn]
            
            for ci, v in enumerate(row_data, 1):
                c = ws.cell(ri,ci,v); c.fill=fill; c.font=bf; c.border=bd
                if report_type in ('stock','full') and ci==6: c.number_format='0.000'

        ws.column_dimensions['B'].width = 34
        fname = f'BSB_Selected_{report_type}_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
        path = os.path.join(folder, fname)
        wb.save(path)
        if messagebox.askyesno('Saved',f'Saved!\n{path}\n\nOpen now?', parent=self):
            os.startfile(path)

    def _qty_dlg(self, reload_fn):
        C = T()
        # Works via ticked checkbox, highlighted row, or double-click
        sel = self._cbt_selection()
        if not sel: messagebox.showinfo('Select','Click an item row (or tick its checkbox) first.',parent=self); return

        multi = len(sel) > 1
        ids = [v[0] for v in sel]
        first = sel[0]
        cur_qty = first[3]

        w = ctk.CTkToplevel(self); w.title('Set Quantity')
        w.geometry('400x300'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,'Set Quantity',C['text'],'h2').pack(pady=(18,4))
        if multi:
            L(w,f'Applying to {len(sel)} selected items',C['gold'],'sm_b').pack(padx=20)
            L(w,'They will all be set to the same quantity',C['text4'],'tiny').pack(pady=(2,16))
        else:
            L(w,first[1],C['text3'],'sm').pack(padx=20)
            L(w,f'Currently in stock: {cur_qty}',C['text4'],'tiny').pack(pady=(2,16))

        qv = tk.StringVar(value=str(cur_qty if not multi else 0))
        row = ctk.CTkFrame(w, fg_color=C['surface']); row.pack()
        def dec(): qv.set(str(max(0,int(qv.get() or 0)-1)))
        def inc(): qv.set(str(int(qv.get() or 0)+1))
        ghost_btn(row,'−',dec,w=44,h=44).pack(side='left',padx=6)
        ctk.CTkEntry(row,textvariable=qv,width=96,height=44,fg_color=C['surface2'],
                    border_color=C['border'],text_color=C['text'],
                    font=ctk.CTkFont('Outfit',22,'bold'),corner_radius=8,justify='center').pack(side='left',padx=6)
        ghost_btn(row,'＋',inc,w=44,h=44).pack(side='left',padx=6)

        def save():
            try:
                nq = int(qv.get())
                for iid in ids:
                    ex("UPDATE Items SET Quantity=? WHERE ItemID=?",(nq,iid))
                reload_fn(); w.destroy()
            except: messagebox.showerror('Invalid','Enter a whole number.',parent=w)
        blue_btn(w,'Save Quantity',save,w=280,h=42).pack(pady=18)
        w.bind('<Return>',lambda e: save())

    def _item_images_dir(self):
        """Folder next to the app where product images are stored."""
        try:
            base = os.path.dirname(os.path.abspath(__file__))
        except:
            base = os.getcwd()
        d = os.path.join(base, 'item_images')
        try: os.makedirs(d, exist_ok=True)
        except: pass
        return d

    def _cbt_selection(self):
        """Return selected item value-tuples. Prefer ticked checkboxes; if none
        are ticked, fall back to the row the user clicked/highlighted."""
        if not hasattr(self, '_cbt'):
            return []
        vals = self._cbt.get_selected_values()
        if vals:
            return vals
        # Fall back to highlighted tree rows
        try:
            sel_items = self._cbt.tree.selection()
        except Exception:
            sel_items = []
        out = []
        for ti in sel_items:
            iid = self._cbt._id_map.get(ti)
            if iid is not None and iid in self._cbt._val_map:
                out.append(self._cbt._val_map[iid])
        return out

    def _show_item_image(self, item_id):
        """Popup showing the product image for an item (or a 'no image' notice)."""
        C = T()
        r = qy("SELECT Name, ISNULL(ImagePath,'') FROM Items WHERE ItemID=?", (item_id,))
        if not r:
            return
        name, img_name = r[0][0], r[0][1]
        w = ctk.CTkToplevel(self); w.title('Item Image')
        w.geometry('460x520'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60, w.lift)
        ctk.CTkFrame(w, height=3, fg_color=C['gold']).pack(fill='x')
        L(w, name, C['text'], 'h2').pack(pady=(18, 4))

        full = os.path.join(self._item_images_dir(), img_name) if img_name else ''
        shown = False
        if img_name and os.path.exists(full):
            try:
                from PIL import Image as PILImage, ImageTk
                pil = PILImage.open(full).convert('RGB')
                pil.thumbnail((400, 400))
                ph = ImageTk.PhotoImage(pil)
                self._view_img_ref = ph  # keep a reference so it isn't GC'd
                holder = ctk.CTkFrame(w, fg_color=C['surface2'], corner_radius=10,
                                      border_width=1, border_color=C['border'])
                holder.pack(padx=24, pady=14)
                tk.Label(holder, image=ph, bg=C['surface2'], bd=0).pack(padx=10, pady=10)
                L(w, img_name, C['text4'], 'tiny').pack()
                shown = True
            except Exception:
                shown = False
        if not shown:
            box = ctk.CTkFrame(w, fg_color=C['surface2'], corner_radius=10,
                               border_width=1, border_color=C['border'])
            box.pack(padx=24, pady=24, fill='both', expand=True)
            inner = ctk.CTkFrame(box, fg_color=C['surface2']); inner.place(relx=0.5, rely=0.5, anchor='center')
            L(inner, '🖼', fk='disp').pack()
            L(inner, 'No image for this item', C['text3'], 'body').pack(pady=(8, 0))
            L(inner, 'Add one via Edit Item → Browse Image', C['text4'], 'tiny').pack(pady=(4, 0))
        blue_btn(w, 'Close', w.destroy, w=200, h=40).pack(pady=(6, 16))

    def _item_dlg(self, reload_fn, edit=False):
        C = T()
        editing = None
        if edit:
            sel = self._cbt_selection()
            if not sel:
                messagebox.showinfo('Select','Click an item row (or tick its checkbox) first, then press Edit.',parent=self); return
            editing = sel[0][0]

        w = ctk.CTkToplevel(self); w.title('Edit Item' if editing else 'Add Item')
        w.geometry('500x760'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,'Edit Item' if editing else 'Add New Item',C['text'],'h2').pack(pady=(18,6))

        # scrollable body so the image row always fits
        fm = ctk.CTkScrollableFrame(w,fg_color=C['surface'],width=430,height=560)
        fm.pack(fill='both',expand=True,padx=24,pady=(0,4))

        existing = None
        if editing:
            r = qy("SELECT Name,Quantity,Unit,UnitPrice,MinStock,SubCatID,ISNULL(ImagePath,''),ISNULL(LPONumber,''),ISNULL(InvoiceNumber,'') FROM Items WHERE ItemID=?",(editing,))
            if r: existing = r[0]

        fields = {}
        for label, key, default in [('Item Name','name',existing[0] if existing else ''),
                                    ('Quantity','qty',str(existing[1]) if existing else '0'),
                                    ('Unit','unit',existing[2] if existing else 'pcs'),
                                    ('Price BHD','price',f'{float(existing[3]):.3f}' if existing else '0.000'),
                                    ('Min Stock','min',str(existing[4]) if existing else '5'),
                                    ('LPO Number (optional)','lpo',(existing[7] if existing else '')),
                                    ('Invoice Number (optional)','invoice',(existing[8] if existing else ''))]:
            caps_lbl(fm, label).pack(anchor='w', pady=(11,3))
            e = inp(fm, w=400); e.insert(0, default); e.pack(); fields[key] = e

        caps_lbl(fm, 'Category').pack(anchor='w', pady=(11,3))
        cats = qy("SELECT s.SubCatID,s.Name FROM SubCategories s ORDER BY s.Name")
        opts = [r[1] for r in cats]; cmap = {r[1]:r[0] for r in cats}; rev = {r[0]:r[1] for r in cats}
        cv = ctk.StringVar(value=(rev.get(existing[5]) if existing and existing[5] in rev else (opts[0] if opts else '')))
        ctk.CTkComboBox(fm, values=opts, variable=cv, width=400, height=40,
                       fg_color=C['surface2'], border_color=C['border'],
                       button_color=C['gold'], button_hover_color=C['gold2'],
                       text_color=C['text'], font=F['body'], corner_radius=6,
                       dropdown_fg_color=C['surface2'], dropdown_text_color=C['text']).pack()

        # ── Product Image (optional) ──
        caps_lbl(fm, 'Product Image  (optional)').pack(anchor='w', pady=(14,3))
        img_state = {
            'src': None,                                   # newly chosen file on disk
            'current': (existing[6] if existing else ''),  # existing stored filename
            'remove': False,                               # user cleared the image
            'preview': None,                               # keep PhotoImage ref
        }
        img_box = ctk.CTkFrame(fm, fg_color=C['surface2'], corner_radius=8,
                               border_width=1, border_color=C['border'])
        img_box.pack(fill='x', pady=(0,4))
        inner = ctk.CTkFrame(img_box, fg_color=C['surface2']); inner.pack(fill='x', padx=12, pady=12)

        # preview thumbnail (left) + filename/buttons (right)
        thumb = tk.Label(inner, bg=C['surface2'], bd=0)
        thumb.pack(side='left', padx=(0,12))
        right = ctk.CTkFrame(inner, fg_color=C['surface2']); right.pack(side='left', fill='x', expand=True)
        fname_lbl = L(right, 'No image selected', C['text3'], 'sm'); fname_lbl.pack(anchor='w')
        btn_row = ctk.CTkFrame(right, fg_color=C['surface2']); btn_row.pack(anchor='w', pady=(8,0))

        def render_preview(path_or_name, on_disk):
            """Show a small thumbnail from a full path (on_disk=True) or stored name."""
            try:
                from PIL import Image as PILImage, ImageTk
                if on_disk:
                    full = path_or_name
                else:
                    full = os.path.join(self._item_images_dir(), path_or_name)
                if not full or not os.path.exists(full):
                    thumb.configure(image='', text=''); return False
                pil = PILImage.open(full).convert('RGB')
                pil.thumbnail((72,72))
                ph = ImageTk.PhotoImage(pil)
                img_state['preview'] = ph
                thumb.configure(image=ph)
                return True
            except Exception:
                thumb.configure(image='', text='')
                return False

        def refresh():
            if img_state['remove'] or (not img_state['src'] and not img_state['current']):
                fname_lbl.configure(text='No image selected', text_color=C['text3'])
                thumb.configure(image='')
                img_state['preview'] = None
                return
            if img_state['src']:
                fname_lbl.configure(text=os.path.basename(img_state['src']), text_color=C['text'])
                render_preview(img_state['src'], True)
            elif img_state['current']:
                fname_lbl.configure(text=img_state['current'], text_color=C['text'])
                render_preview(img_state['current'], False)

        def browse():
            p = filedialog.askopenfilename(
                title='Choose product image', parent=w,
                filetypes=[('Image files','*.png *.jpg *.jpeg *.gif *.bmp *.webp'),
                           ('All files','*.*')])
            if p:
                img_state['src'] = p
                img_state['remove'] = False
                refresh()

        def clear_img():
            img_state['src'] = None
            img_state['remove'] = True
            refresh()

        gold_btn(btn_row, '📁 Browse Image', browse, w=150, h=34).pack(side='left', padx=(0,8))
        ghost_btn(btn_row, '✕ Remove', clear_img, color=C['red'], w=100, h=34).pack(side='left')
        L(fm, 'PNG / JPG / GIF · stored in the item_images folder', C['text4'], 'tiny').pack(anchor='w', pady=(6,0))
        refresh()

        def save():
            try:
                n = fields['name'].get().strip()
                if not n: raise ValueError('Name required')

                # Resolve final image filename
                image_name = img_state['current']  # keep existing by default
                if img_state['remove']:
                    image_name = ''
                if img_state['src']:
                    # copy the chosen file into item_images with a safe unique name
                    import shutil, time as _t
                    ext = os.path.splitext(img_state['src'])[1].lower() or '.png'
                    safe = ''.join(ch for ch in n if ch.isalnum() or ch in (' ','-','_')).strip().replace(' ','_')[:40]
                    image_name = f"{safe}_{int(_t.time())}{ext}"
                    try:
                        shutil.copy(img_state['src'], os.path.join(self._item_images_dir(), image_name))
                    except Exception as ce:
                        messagebox.showerror('Image error', f'Could not save image:\n{ce}', parent=w); return

                args = (n, cmap.get(cv.get()), int(fields['qty'].get()),
                        fields['unit'].get().strip(), float(fields['price'].get()), int(fields['min'].get()),
                        image_name, fields['lpo'].get().strip(), fields['invoice'].get().strip())
                if editing:
                    ex("UPDATE Items SET Name=?,SubCatID=?,Quantity=?,Unit=?,UnitPrice=?,MinStock=?,ImagePath=?,LPONumber=?,InvoiceNumber=? WHERE ItemID=?",args+(editing,))
                else:
                    ex("INSERT INTO Items(Name,SubCatID,Quantity,Unit,UnitPrice,MinStock,ImagePath,LPONumber,InvoiceNumber) VALUES(?,?,?,?,?,?,?,?,?)",args)
                reload_fn(); w.destroy()
            except Exception as e:
                messagebox.showerror('Error',str(e),parent=w)
        blue_btn(w,'Save Item',save,w=428,h=42).pack(pady=(4,16),padx=36)

    def _del_items(self, reload_fn):
        sel = self._cbt_selection()
        if not sel: messagebox.showinfo('Select','Click an item row (or tick its checkbox) first.',parent=self); return
        C = T()
        w = ctk.CTkToplevel(self); w.title('Confirm Delete')
        w.geometry('380x200'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['red']).pack(fill='x')
        L(w,'Delete items?',C['text'],'h2').pack(pady=(18,8))
        L(w,f'{len(sel)} item(s) will be permanently deleted.',C['text3'],'sm').pack()
        bf = ctk.CTkFrame(w,fg_color=C['surface']); bf.pack(pady=20)
        ghost_btn(bf,'Cancel',w.destroy,w=110,h=38).pack(side='left',padx=8)
        def confirm():
            for r in sel: ex("DELETE FROM Items WHERE ItemID=?",(r[0],))
            reload_fn(); w.destroy()
        red_btn(bf,'Delete All',confirm,w=120,h=38).pack(side='left',padx=8)

    def _tab_cats(self, parent):
        C = T()
        main = ctk.CTkFrame(parent,fg_color=C['bg']); main.pack(fill='both',expand=True,padx=16,pady=14)
        def panel(title):
            p = ctk.CTkFrame(main,fg_color=C['card'],corner_radius=12,border_width=1,border_color=C['border'])
            p.pack(side='left',fill='both',expand=True,padx=7)
            L(p,title,C['text'],'h3').pack(pady=14)
            hline(p)
            holder_box = ctk.CTkFrame(p,fg_color=C['card']); holder_box.pack(fill='both',expand=True,padx=10,pady=8)
            hline(p)
            bf = ctk.CTkFrame(p,fg_color=C['card']); bf.pack(pady=10)
            return holder_box, bf

        pbox, pb = panel('Parent Categories')
        sbox, sb = panel('Sub-Categories')

        pcbt = CheckboxTree(pbox, ('Category','Items'), [220,80], height=14, simple_mode=True)
        pcbt.frame.pack(fill='both', expand=True)
        scbt = CheckboxTree(sbox, ('Sub-Category','Parent','Items'), [170,140,70], height=14, simple_mode=True)
        scbt.frame.pack(fill='both', expand=True)

        def lp():
            pcbt.clear()
            for r in qy("""SELECT p.ParentCatID, p.Name,
                                  (SELECT COUNT(*) FROM Items i JOIN SubCategories s ON i.SubCatID=s.SubCatID WHERE s.ParentCatID=p.ParentCatID)
                           FROM ParentCategories p ORDER BY p.Name"""):
                pcbt.insert(r[0], (r[1], r[2]))
        def ls():
            scbt.clear()
            for r in qy("""SELECT s.SubCatID, s.Name, p.Name,
                                  (SELECT COUNT(*) FROM Items i WHERE i.SubCatID=s.SubCatID)
                           FROM SubCategories s JOIN ParentCategories p ON s.ParentCatID=p.ParentCatID
                           ORDER BY p.Name, s.Name"""):
                scbt.insert(r[0], (r[1], r[2], r[3]))

        def ap():
            n = self._ask('Add Parent Category','Category name:')
            if n: ex("INSERT INTO ParentCategories(Name) VALUES(?)",(n,)); lp()
        def dp():
            ids = list(pcbt.checked) or pcbt.all_ids[:0]
            if not ids:
                # fall back to highlighted row if nothing is ticked
                sel = pcbt.tree.selection()
                if sel:
                    iid = pcbt._id_map.get(sel[0])
                    if iid is not None: ids=[iid]
            if not ids:
                messagebox.showinfo('Select','Tick the categories you want to delete (or click a row).',parent=self); return
            n=len(ids)
            if messagebox.askyesno('Delete',f'Delete {n} categor{"y" if n==1 else "ies"} and all their sub-categories?',parent=self):
                for pid in ids: ex("DELETE FROM ParentCategories WHERE ParentCatID=?",(pid,))
                lp(); ls()
        def as_():
            parents=qy("SELECT ParentCatID,Name FROM ParentCategories ORDER BY Name")
            if not parents: return
            w=ctk.CTkToplevel(self); w.title('Add Sub-Category')
            w.geometry('360x250'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
            ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
            L(w,'Add Sub-Category',C['text'],'h2').pack(pady=(16,8))
            fm=ctk.CTkFrame(w,fg_color=C['surface']); fm.pack(fill='x',padx=30)
            caps_lbl(fm,'Name').pack(anchor='w',pady=(8,3)); ne=inp(fm,w=300); ne.pack()
            caps_lbl(fm,'Parent').pack(anchor='w',pady=(10,3))
            pmap={r[1]:r[0] for r in parents}; pv=ctk.StringVar(value=list(pmap.keys())[0])
            ctk.CTkComboBox(fm,values=list(pmap.keys()),variable=pv,width=300,height=38,
                           fg_color=C['surface2'],border_color=C['border'],button_color=C['gold'],
                           text_color=C['text'],font=F['body'],corner_radius=6,
                           dropdown_fg_color=C['surface2'],dropdown_text_color=C['text']).pack()
            def sv(): n=ne.get().strip(); n and ex("INSERT INTO SubCategories(Name,ParentCatID) VALUES(?,?)",(n,pmap.get(pv.get()))) and ls() or ls() and w.destroy()
            blue_btn(w,'Save',sv,w=300,h=38).pack(pady=14)
        def ds():
            ids = list(scbt.checked)
            if not ids:
                sel = scbt.tree.selection()
                if sel:
                    iid = scbt._id_map.get(sel[0])
                    if iid is not None: ids=[iid]
            if not ids:
                messagebox.showinfo('Select','Tick the sub-categories you want to delete (or click a row).',parent=self); return
            n=len(ids)
            if messagebox.askyesno('Delete',f'Delete {n} sub-categor{"y" if n==1 else "ies"}?',parent=self):
                for sid in ids: ex("DELETE FROM SubCategories WHERE SubCatID=?",(sid,))
                ls()
        blue_btn(pb,'+ Add',ap,w=110,h=36).pack(side='left',padx=4)
        red_btn(pb,'🗑 Del Selected',dp,w=140,h=36).pack(side='left',padx=4)
        blue_btn(sb,'+ Add',as_,w=110,h=36).pack(side='left',padx=4)
        red_btn(sb,'🗑 Del Selected',ds,w=140,h=36).pack(side='left',padx=4)
        lp(); ls()

    def _ask(self, title, label):
        C = T()
        w=ctk.CTkToplevel(self); w.title(title)
        w.geometry('360x180'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,title,C['text'],'h2').pack(pady=(16,8))
        e=inp(w,w=300); e.pack(padx=30); e.focus()
        result=[None]
        def ok(): result[0]=e.get().strip(); w.destroy()
        blue_btn(w,'Save',ok,w=300,h=38).pack(pady=14)
        w.wait_window(); return result[0]

    # ══ ISSUE ════════════════════════════════════════════════
    def _order(self):
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root=ctk.CTkFrame(self,fg_color=C['bg']); root.pack(fill='both',expand=True)
        self._topbar(root,'Stationery','Place an Order',lambda: self._reload(self._order))
        body=ctk.CTkFrame(root,fg_color=C['bg']); body.pack(fill='both',expand=True)
        filt={'p':None,'s':None}
        def on_filter(p,s): filt['p']=p; filt['s']=s; load()
        self._sidebar(body,on_filter)

        mid=ctk.CTkFrame(body,fg_color=C['bg']); mid.pack(side='left',fill='both',expand=True)
        tb=ctk.CTkFrame(mid,fg_color=C['topbar'],corner_radius=0,height=58)
        tb.pack(fill='x'); tb.pack_propagate(False)
        sv=ctk.StringVar()
        L(tb,'Search',C['text4'],'xs').pack(side='left',padx=(16,8),pady=14)
        ctk.CTkEntry(tb,textvariable=sv,placeholder_text='Search items...',width=280,height=34,
                    fg_color=C['surface2'],border_color=C['border'],text_color=C['text'],
                    placeholder_text_color=C['text4'],font=F['body'],corner_radius=6).pack(side='left',pady=14)
        sv.trace_add('write',lambda *_: load())
        blue_btn(tb,'🛒  Review & Order Selected',lambda: review(),w=220,h=36).pack(side='right',padx=10,pady=12)

        hint=ctk.CTkFrame(mid,fg_color=C['bg'],height=26); hint.pack(fill='x'); hint.pack_propagate(False)
        L(hint,'ⓘ  Click "🖼 View" to see an item · Tick the items you want, then click "Review & Order Selected".',C['text4'],'tiny').pack(side='left',padx=18,pady=4)

        tcard=ctk.CTkFrame(mid,fg_color=C['card'],corner_radius=12,border_width=1,border_color=C['border'])
        tcard.pack(fill='both',expand=True,padx=12,pady=(2,12))
        # Staff do NOT see stock — only ID, Item Name, Category, Unit (+ image button)
        self._ocbt=CheckboxTree(tcard,('ID','Item Name','Category','Unit'),[60,360,220,140],height=18,order_mode=True,image_col=True)
        self._ocbt.frame.pack(fill='both',expand=True,padx=2,pady=2)
        self._ocbt.on_image_click = lambda item_id: self._show_item_image(item_id)

        def load():
            self._ocbt.clear()
            base="SELECT i.ItemID,i.Name,ISNULL(s.Name,'-'),i.Unit FROM Items i LEFT JOIN SubCategories s ON i.SubCatID=s.SubCatID "
            srch=f"%{sv.get()}%"
            if filt['s']:   rows=qy(base+"WHERE i.SubCatID=? AND i.Name LIKE ? ORDER BY i.Name",(filt['s'],srch))
            elif filt['p']: rows=qy(base+"WHERE s.ParentCatID=? AND i.Name LIKE ? ORDER BY s.Name,i.Name",(filt['p'],srch))
            else:           rows=qy(base+"WHERE i.Name LIKE ? ORDER BY s.Name,i.Name",(srch,))
            for r in rows:
                self._ocbt.insert(r[0],(r[0],r[1],r[2],r[3]))
        load()

        def review():
            sel=self._ocbt.get_selected_values()
            if not sel: messagebox.showinfo('Select','Tick the items you want to order first.',parent=self); return
            self._order_review_dlg(sel)

    def _order_review_dlg(self, sel):
        C = T()
        w=ctk.CTkToplevel(self); w.title('Review Order')
        w.geometry('560x620'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,'Review Your Order',C['text'],'h2').pack(pady=(18,2))
        L(w,'Set the quantity for each item, then submit.',C['text3'],'sm').pack(pady=(0,10))

        # Scrollable list of selected items with qty spinners
        sc=ctk.CTkScrollableFrame(w,fg_color=C['surface'],width=500,height=320)
        sc.pack(padx=20,pady=(0,10),fill='both',expand=True)
        qty_vars={}
        for v in sel:
            iid,name,cat,unit = v
            rowf=ctk.CTkFrame(sc,fg_color=C['card'],corner_radius=8,border_width=1,border_color=C['border'])
            rowf.pack(fill='x',pady=4)
            inner=ctk.CTkFrame(rowf,fg_color=C['card']); inner.pack(fill='x',padx=12,pady=8)
            tx=ctk.CTkFrame(inner,fg_color=C['card']); tx.pack(side='left',fill='x',expand=True)
            L(tx,name,C['text'],'sm_b').pack(anchor='w')
            L(tx,f'{cat}  ·  {unit}',C['text4'],'tiny').pack(anchor='w')
            qv=tk.StringVar(value='1'); qty_vars[iid]=(qv,name,unit)
            sp=ctk.CTkFrame(inner,fg_color=C['card']); sp.pack(side='right')
            def mk(qv):
                def dec(): qv.set(str(max(1,int(qv.get() or 1)-1)))
                def inc(): qv.set(str(int(qv.get() or 0)+1))
                return dec,inc
            dec,inc=mk(qv)
            ghost_btn(sp,'−',dec,w=36,h=34).pack(side='left',padx=3)
            ctk.CTkEntry(sp,textvariable=qv,width=64,height=34,fg_color=C['surface2'],
                        border_color=C['border'],text_color=C['text'],
                        font=ctk.CTkFont('Outfit',16,'bold'),corner_radius=6,justify='center').pack(side='left',padx=3)
            ghost_btn(sp,'＋',inc,w=36,h=34).pack(side='left',padx=3)

        # Department + notes + submit
        fm=ctk.CTkFrame(w,fg_color=C['surface']); fm.pack(fill='x',padx=20)
        caps_lbl(fm,'Department / Class  (required)').pack(anchor='w',pady=(4,3)); de=inp(fm,w=500); de.pack()
        derr=ctk.CTkLabel(fm,text='',text_color=C['red'],font=F['tiny']); derr.pack(anchor='w')
        caps_lbl(fm,'Notes (optional)').pack(anchor='w',pady=(4,3)); no=inp(fm,w=500); no.pack()
        def submit():
            dept=de.get().strip()
            if not dept:
                derr.configure(text='✕  Department / Class is required.')
                return
            items=[]
            for iid,(qv,name,unit) in qty_vars.items():
                try: q=int(qv.get())
                except: q=1
                if q>0: items.append((iid,name,q,unit))
            if not items: messagebox.showerror('Empty','Set at least one quantity.',parent=w); return
            ex("INSERT INTO Orders(OrderedBy,OrderedByName,Department,Status,Notes,OrderedByEmail) VALUES(?,?,?,?,?,?)",
               (self.current_username,self.current_user,dept,'Pending',no.get(),self.current_email))
            oid=qy("SELECT TOP 1 OrderID FROM Orders ORDER BY OrderID DESC")[0][0]
            for iid,name,q,unit in items:
                ex("INSERT INTO OrderItems(OrderID,ItemID,ItemName,Quantity,Unit) VALUES(?,?,?,?,?)",
                   (oid,iid,name,q,unit))
            w.destroy()
            messagebox.showinfo('Order Sent','Your order was sent to the admin for arranging!',parent=self)
            self._order()
        blue_btn(w,'✔  Send Order to Admin',submit,w=500,h=44).pack(pady=14,padx=20)

    # ══ STAFF: ORDER HISTORY ═════════════════════════════════
    def _order_history(self):
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root=ctk.CTkFrame(self,fg_color=C['bg']); root.pack(fill='both',expand=True)
        self._topbar(root,'Stationery','Orders History',lambda: self._reload(self._order_history))
        sc=ctk.CTkScrollableFrame(root,fg_color=C['bg']); sc.pack(fill='both',expand=True,padx=24,pady=18)
        L(sc,'Your Order History',C['text'],'h2').pack(anchor='w',pady=(0,14))

        orders=qy("""SELECT OrderID,Department,Status,CreatedAt,ISNULL(Notes,'')
                     FROM Orders WHERE OrderedBy=? ORDER BY CreatedAt DESC""",(self.current_username,))
        if not orders:
            L(sc,'You have not placed any orders yet.',C['text3'],'body').pack(anchor='w',pady=20)
            return
        status_colors={'Pending':(C['low_bg'],C['low_fg']),'Arranged':(C['in_bg'],C['in_fg']),
                       'Collected':(C['surface2'],C['text3'])}
        for oid,dept,status,created,notes in orders:
            card=ctk.CTkFrame(sc,fg_color=C['card'],corner_radius=10,border_width=1,border_color=C['border'])
            card.pack(fill='x',pady=6)
            head=ctk.CTkFrame(card,fg_color=C['card']); head.pack(fill='x',padx=16,pady=(12,4))
            dt=created.strftime('%d/%m/%Y %H:%M') if hasattr(created,'strftime') else str(created)
            L(head,f'Order #{oid}',C['text'],'h3').pack(side='left')
            sbg,sfg=status_colors.get(status,(C['surface2'],C['text3']))
            ctk.CTkLabel(head,text=f' {status} ',text_color=sfg,fg_color=sbg,font=F['caps'],corner_radius=8).pack(side='right')
            L(head,dt,C['text4'],'tiny').pack(side='right',padx=10)
            if dept: L(card,f'Department: {dept}',C['text3'],'sm').pack(anchor='w',padx=16)
            items=qy("SELECT ItemName,Quantity,ISNULL(Unit,'') FROM OrderItems WHERE OrderID=?",(oid,))
            itxt='  •  '.join(f"{q}× {nm}" for nm,q,u in items)
            L(card,itxt,C['text2'],'sm').pack(anchor='w',padx=16,pady=(2,12))

    # ══ ADMIN: ORDERS / MAIL ═════════════════════════════════
    def _orders_mail(self, order_filter=None):
        # order_filter: None/'all' = all, else 'Pending','Arranged','Collected','OutOfStock','New','Date'
        if order_filter is None:
            order_filter = getattr(self, '_order_filter', 'all')
        self._order_filter = order_filter
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root=ctk.CTkFrame(self,fg_color=C['bg']); root.pack(fill='both',expand=True)
        self._topbar(root,'Procurement','Orders',lambda: self._reload(self._orders_mail))
        body=ctk.CTkFrame(root,fg_color=C['bg']); body.pack(fill='both',expand=True)

        # ── Left stock-check panel ────────────────────────
        panel=tk.Frame(body,bg=C['sidebar'],width=300); panel.pack(side='left',fill='y')
        panel.pack_propagate(False)
        tk.Frame(panel,bg='#1A3A5C',width=1).pack(side='right',fill='y')
        pw=tk.Frame(panel,bg=C['sidebar']); pw.pack(fill='both',expand=True)
        tk.Label(pw,text='STOCK CHECK',bg=C['sidebar'],fg=C['side_mut'],
                font=('Outfit',9,'bold'),anchor='w',padx=16,pady=14).pack(fill='x')
        tk.Frame(pw,bg='#1A3A5C',height=1).pack(fill='x')

        # Build stock check: for each pending order, check if items are in stock
        pcanvas=tk.Canvas(pw,bg=C['sidebar'],highlightthickness=0,bd=0)
        pscr=tk.Scrollbar(pw,orient='vertical',command=pcanvas.yview)
        pnav=tk.Frame(pcanvas,bg=C['sidebar'])
        pnav.bind('<Configure>',lambda e: pcanvas.configure(scrollregion=pcanvas.bbox('all')))
        pcanvas.create_window((0,0),window=pnav,anchor='nw',width=284)
        pcanvas.configure(yscrollcommand=pscr.set)
        pcanvas.pack(side='left',fill='both',expand=True)
        pscr.pack(side='right',fill='y')

        pending_orders=qy("""SELECT o.OrderID,o.OrderedByName
                             FROM Orders o WHERE o.Status='Pending'
                             ORDER BY o.CreatedAt""")
        for oid,who in pending_orders:
            items=qy("""SELECT oi.ItemName,oi.Quantity,ISNULL(i.Quantity,0)
                        FROM OrderItems oi
                        LEFT JOIN Items i ON oi.ItemID=i.ItemID
                        WHERE oi.OrderID=?""",(oid,))
            all_ok = all(avail >= req for _,req,avail in items)
            hdr_bg = C['sidebar'] if all_ok else C['out_bg']
            hdr_fg = C['gold'] if all_ok else C['out_fg']
            of=tk.Frame(pnav,bg=hdr_bg); of.pack(fill='x',pady=2,padx=6)
            tk.Label(of,text=f'Order #{oid}',bg=hdr_bg,fg=hdr_fg,
                    font=('Outfit',12,'bold'),anchor='w',padx=10,pady=6).pack(fill='x')
            tk.Label(of,text=who,bg=hdr_bg,fg=C['side_mut'],
                    font=('Outfit',10),anchor='w',padx=10).pack(fill='x')
            for nm,req,avail in items:
                ok = avail >= req
                ic = C['in_fg'] if ok else C['out_fg']
                txt = f"  {'✓' if ok else '✕'}  {nm[:22]}  ({req} needed, {avail} avail)"
                tk.Label(of,text=txt,bg=hdr_bg,fg=ic,
                        font=('Outfit',9),anchor='w',padx=10,pady=2,wraplength=270,justify='left').pack(fill='x')

        # ── Right orders scroll ───────────────────────────
        sc=ctk.CTkScrollableFrame(body,fg_color=C['bg'])
        sc.pack(side='left',fill='both',expand=True,padx=20,pady=18)
        L(sc,'🛒  Staff Orders',C['text'],'h2').pack(anchor='w',pady=(0,4))
        L(sc,'Check stock panel on the left. Mark Arranged deducts stock.',C['text3'],'sm').pack(anchor='w',pady=(0,10))

        # ── Filter bar ─────────────────────────────────────
        filt_wrap = ctk.CTkFrame(sc, fg_color=C['bg']); filt_wrap.pack(fill='x', pady=(0,4))
        L(filt_wrap, 'Filter:', C['text4'], 'xs').pack(side='left', padx=(0,8))
        filter_opts = [
            ('All', 'all'),
            ('New Orders', 'New'),
            ('Pending', 'Pending'),
            ('Not Collected', 'NotCollected'),
            ('Collected', 'Collected'),
            ('By Date', 'Date'),
        ]
        for label, key in filter_opts:
            active = (self._order_filter == key) or (self._order_filter in (None,'all') and key=='all')
            b = ctk.CTkButton(filt_wrap, text=label,
                              command=lambda k=key: self._orders_mail(k),
                              fg_color=C['blue'] if active else C['surface2'],
                              hover_color='#1746B0' if active else C['border'],
                              text_color='white' if active else C['text3'],
                              font=F['caps'], corner_radius=14, height=28, width=96,
                              border_width=0 if active else 1, border_color=C['border'])
            b.pack(side='left', padx=3)

        # Orders report button
        blue_btn(sc,'📊 Export Orders Report',lambda: self._e_orders(),w=220,h=36).pack(anchor='w',pady=(8,14))

        # Build the query depending on the chosen filter
        flt = self._order_filter
        if flt == 'Pending':
            where = "WHERE Status='Pending'"; order = "CreatedAt DESC"
        elif flt == 'New':
            # New = pending orders, newest first
            where = "WHERE Status='Pending'"; order = "CreatedAt DESC"
        elif flt == 'NotCollected':
            where = "WHERE Status IN ('Pending','Arranged','OutOfStock')"; order = "CreatedAt DESC"
        elif flt == 'Collected':
            where = "WHERE Status='Collected'"; order = "CreatedAt DESC"
        elif flt == 'Date':
            where = ""; order = "CreatedAt DESC"
        else:  # all
            where = ""
            order = ("CASE Status WHEN 'Pending' THEN 0 WHEN 'Arranged' THEN 1 "
                     "WHEN 'OutOfStock' THEN 2 ELSE 3 END, CreatedAt DESC")

        orders=qy(f"""SELECT OrderID,OrderedByName,Department,Status,CreatedAt,ISNULL(Notes,''),ISNULL(OrderedByEmail,'')
                     FROM Orders {where} ORDER BY {order}""")
        if not orders:
            L(sc,'No orders match this filter.',C['text3'],'body').pack(anchor='w',pady=10)
        status_colors={'Pending':(C['low_bg'],C['low_fg']),'Arranged':(C['in_bg'],C['in_fg']),
                       'Collected':(C['surface2'],C['text3']),'OutOfStock':(C['out_bg'],C['out_fg'])}

        for oid,who,dept,status,created,notes,email in orders:
            # Check if any items are out of stock (for border color)
            items=qy("""SELECT oi.ItemName,oi.Quantity,ISNULL(i.Quantity,0),ISNULL(i.Unit,'')
                        FROM OrderItems oi LEFT JOIN Items i ON oi.ItemID=i.ItemID
                        WHERE oi.OrderID=?""",(oid,))
            has_stock_issue = status=='Pending' and any(avail < req for _,req,avail,_ in items)
            bd_color = C['out_fg'] if has_stock_issue else C['border']

            card=ctk.CTkFrame(sc,fg_color=C['card'],corner_radius=10,border_width=2 if has_stock_issue else 1,border_color=bd_color)
            card.pack(fill='x',pady=6)
            head=ctk.CTkFrame(card,fg_color=C['card']); head.pack(fill='x',padx=16,pady=(12,4))
            dt=created.strftime('%d/%m/%Y %H:%M') if hasattr(created,'strftime') else str(created)
            L(head,f'Order #{oid}  —  {who}',C['text'],'h3').pack(side='left')
            sbg,sfg=status_colors.get(status,(C['surface2'],C['text3']))
            ctk.CTkLabel(head,text=f' {status} ',text_color=sfg,fg_color=sbg,font=F['caps'],corner_radius=8).pack(side='right')
            L(head,dt,C['text4'],'tiny').pack(side='right',padx=10)
            meta=f'Department: {dept}' if dept else ''
            if notes: meta+=(f'    ·    Note: {notes}' if meta else f'Note: {notes}')
            if meta: L(card,meta,C['text3'],'sm').pack(anchor='w',padx=16)

            # Items with bold qty
            irow=ctk.CTkFrame(card,fg_color=C['card']); irow.pack(anchor='w',padx=16,pady=(2,8))
            for nm,req,avail,u in items:
                ok=avail>=req
                ctk.CTkLabel(irow,text=f'{req}×',text_color=C['out_fg'] if not ok else C['gold'],
                            font=ctk.CTkFont('Outfit',16,'bold')).pack(side='left',padx=(0,4))
                ctk.CTkLabel(irow,text=f'{nm}',text_color=C['out_fg'] if not ok else C['text'],
                            font=ctk.CTkFont('Outfit',13)).pack(side='left',padx=(0,8))
                if not ok and status=='Pending':
                    ctk.CTkLabel(irow,text=f'(only {avail} in stock)',text_color=C['out_fg'],
                                font=ctk.CTkFont('Outfit',10,'bold')).pack(side='left',padx=(0,10))

            bf=ctk.CTkFrame(card,fg_color=C['card']); bf.pack(anchor='w',padx=12,pady=(0,12))

            def arrange(o=oid,it=items,em=email,wh=who):
                # Check stock first
                short=[nm for nm,req,avail,u in it if avail<req]
                if short:
                    if not messagebox.askyesno('Stock Warning',
                        f'Some items are low/out of stock:\n{chr(10).join(short)}\n\nMark as Arranged anyway?',parent=self):
                        return
                # Deduct stock
                for nm,req,avail,u in it:
                    ex("""UPDATE Items SET Quantity=Quantity-? 
                          WHERE ItemID=(SELECT TOP 1 ItemID FROM Items WHERE Name=?)
                          AND Quantity>=?""",(req,nm,req))
                ex("UPDATE Orders SET Status='Arranged' WHERE OrderID=?",(o,))
                # Email with actual items
                if em and SMTP_CONFIG.get('enabled'):
                    try:
                        import smtplib; from email.mime.text import MIMEText
                        item_lines=chr(10).join(f"  • {r}× {n}" for n,r,a,u in it)
                        body_txt=(f"Dear {wh},\n\nYour stationery order is ready for collection:\n\n"
                                 f"{item_lines}\n\n"
                                 f"Please come to the Procurement office to collect your items.\n\n"
                                 f"BSB Resource Manager")
                        msg=MIMEText(body_txt); msg['Subject']='BSB Resource Manager — Order Ready'
                        msg['From']=SMTP_CONFIG['sender']; msg['To']=em
                        s2=smtplib.SMTP(SMTP_CONFIG['host'],SMTP_CONFIG['port'],timeout=15)
                        s2.starttls(); s2.login(SMTP_CONFIG['sender'],SMTP_CONFIG['password'])
                        s2.sendmail(SMTP_CONFIG['sender'],[em],msg.as_string()); s2.quit()
                    except Exception as e3: print(f"Email err: {e3}")
                self._orders_mail()

            def out_of_stock(o=oid,em=email,wh=who):
                if messagebox.askyesno('Confirm',f'Mark Order #{o} as Out of Stock and notify {wh}?',parent=self):
                    ex("UPDATE Orders SET Status='OutOfStock' WHERE OrderID=?",(o,))
                    if em and SMTP_CONFIG.get('enabled'):
                        try:
                            import smtplib; from email.mime.text import MIMEText
                            body_txt=(f"Dear {wh},\n\nYour stationery order #{o} is currently out of stock.\n"
                                     f"We will remind you once it becomes available.\n\n"
                                     f"Thank you,\nStationery Procurement Team")
                            msg=MIMEText(body_txt); msg['Subject']='BSB Resource Manager — Order Update'
                            msg['From']=SMTP_CONFIG['sender']; msg['To']=em
                            s2=smtplib.SMTP(SMTP_CONFIG['host'],SMTP_CONFIG['port'],timeout=15)
                            s2.starttls(); s2.login(SMTP_CONFIG['sender'],SMTP_CONFIG['password'])
                            s2.sendmail(SMTP_CONFIG['sender'],[em],msg.as_string()); s2.quit()
                        except Exception as e3: print(f"Email err: {e3}")
                    self._orders_mail()

            if status=='Pending':
                blue_btn(bf,'✔ Mark Arranged',arrange,w=150,h=34).pack(side='left',padx=4)
                ghost_btn(bf,'✕ Out of Stock',out_of_stock,color=C['out_fg'],w=130,h=34).pack(side='left',padx=4)
            elif status=='Arranged':
                gold_btn(bf,'📦 Mark Collected',lambda o=oid: (ex("UPDATE Orders SET Status='Collected' WHERE OrderID=?",(o,)), self._orders_mail()),w=160,h=34).pack(side='left',padx=4)
            elif status=='OutOfStock':
                blue_btn(bf,'↩ Back to Pending',lambda o=oid: (ex("UPDATE Orders SET Status='Pending' WHERE OrderID=?",(o,)), self._orders_mail()),w=160,h=34).pack(side='left',padx=4)


    # ══ ADMIN: ISSUE LOG (tick off completed) ════════════════
    def _issue_log(self):
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root=ctk.CTkFrame(self,fg_color=C['bg']); root.pack(fill='both',expand=True)
        self._topbar(root,'Procurement','Issue Log',lambda: self._reload(self._issue_log))
        sc=ctk.CTkScrollableFrame(root,fg_color=C['bg']); sc.pack(fill='both',expand=True,padx=24,pady=18)
        L(sc,'Issue Log — Remaining & Completed',C['text'],'h2').pack(anchor='w',pady=(0,4))
        L(sc,'Tick off an issue once the items have been handed over.',C['text3'],'sm').pack(anchor='w',pady=(0,16))

        issues=qy("""SELECT r.IssueID,r.RecipientName,r.Department,r.IssuedBy,r.IssuedAt,ISNULL(r.Completed,0)
                     FROM IssueRecords r ORDER BY ISNULL(r.Completed,0), r.IssuedAt DESC""")
        if not issues:
            L(sc,'No issues recorded yet.',C['text3'],'body').pack(anchor='w',pady=10)
        for iid,recip,dept,by,at,done in issues:
            done=bool(done)
            card=ctk.CTkFrame(sc,fg_color=C['surface2'] if done else C['card'],corner_radius=10,
                             border_width=1,border_color=C['border'])
            card.pack(fill='x',pady=5)
            head=ctk.CTkFrame(card,fg_color=card.cget('fg_color')); head.pack(fill='x',padx=16,pady=(10,2))
            dt=at.strftime('%d/%m/%Y %H:%M') if hasattr(at,'strftime') else str(at)
            title=f'Issue #{iid}  —  {recip}'
            L(head,title,C['text3'] if done else C['text'],'h3').pack(side='left')
            if done:
                ctk.CTkLabel(head,text=' ✓ Completed ',text_color=C['in_fg'],fg_color=C['in_bg'],
                            font=F['caps'],corner_radius=8).pack(side='right')
            else:
                ctk.CTkLabel(head,text=' Remaining ',text_color=C['low_fg'],fg_color=C['low_bg'],
                            font=F['caps'],corner_radius=8).pack(side='right')
            L(head,f'by {by} · {dt}',C['text4'],'tiny').pack(side='right',padx=10)
            if dept: L(card,f'Department: {dept}',C['text3'],'sm').pack(anchor='w',padx=16)
            items=qy("SELECT ItemName,Quantity FROM IssueItems WHERE IssueID=?",(iid,))
            itxt='  •  '.join(f"{q}× {nm}" for nm,q in items)
            L(card,itxt,C['text2'],'sm').pack(anchor='w',padx=16,pady=(2,8))
            bf=ctk.CTkFrame(card,fg_color=card.cget('fg_color')); bf.pack(anchor='w',padx=12,pady=(0,12))
            def toggle(i,d):
                ex("UPDATE IssueRecords SET Completed=? WHERE IssueID=?",(0 if d else 1,i)); self._issue_log()
            if done:
                ghost_btn(bf,'↩ Mark Remaining',lambda i=iid,d=done: toggle(i,d),w=150,h=34).pack(side='left',padx=4)
            else:
                blue_btn(bf,'✔ Tick Off (Done)',lambda i=iid,d=done: toggle(i,d),w=160,h=34).pack(side='left',padx=4)

    # ══ DESKTOP NOTIFICATIONS (admin) ════════════════════════
    def _start_notifications(self):
        # seed seen sets so old items don't pop on first run
        try:
            for r in qy("SELECT OrderID FROM Orders WHERE Status='Pending'"): self._notif_seen_orders.add(r[0])
        except: pass
        self._poll_notifications()

    def _poll_notifications(self):
        if self.current_role not in PROCUREMENT_ROLES:
            return
        try:
            for oid,who in qy("SELECT OrderID,OrderedByName FROM Orders WHERE Status='Pending'"):
                if oid not in self._notif_seen_orders:
                    self._notif_seen_orders.add(oid)
                    self._show_notification('🛒 New Order', f'{who} placed order #{oid}')
        except: pass
        # poll again in 8 seconds
        self.after(8000, self._poll_notifications)

    def _show_notification(self, title, body):
        C = T()
        try:
            toast = tk.Toplevel(self)
            toast.overrideredirect(True)
            toast.attributes('-topmost', True)
            w, h = 320, 90
            sw = toast.winfo_screenwidth(); sh = toast.winfo_screenheight()
            x = sw - w - 24; y = sh - h - 60
            toast.geometry(f'{w}x{h}+{x}+{y}')
            frame = tk.Frame(toast, bg='#0F2540', highlightbackground=C['gold'], highlightthickness=2)
            frame.pack(fill='both', expand=True)
            tk.Label(frame, text=title, bg='#0F2540', fg=C['gold'],
                    font=('Outfit', 12, 'bold'), anchor='w').pack(fill='x', padx=14, pady=(10,0))
            tk.Label(frame, text=body[:48], bg='#0F2540', fg='white',
                    font=('Outfit', 11), anchor='w').pack(fill='x', padx=14)
            tk.Label(frame, text='Click to open Orders →', bg='#0F2540', fg='#8892A4',
                    font=('Outfit', 9), anchor='w').pack(fill='x', padx=14, pady=(2,8))
            def open_om(e=None):
                try: toast.destroy()
                except: pass
                self._orders_mail()
            for wdg in [frame]+list(frame.winfo_children()):
                wdg.bind('<Button-1>', open_om)
                wdg.configure(cursor='hand2')
            toast.after(6000, lambda: (toast.destroy() if toast.winfo_exists() else None))
        except Exception:
            pass


    # ══ ADMIN: USER ADMINISTRATION ═══════════════════════════
    def _user_admin(self):
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root=ctk.CTkFrame(self,fg_color=C['bg']); root.pack(fill='both',expand=True)
        self._topbar(root,'Accounts','User Administration',lambda: self._reload(self._user_admin))

        # Toolbar
        tbrow=ctk.CTkFrame(root,fg_color=C['bg']); tbrow.pack(fill='x',padx=24,pady=(16,8))
        L(tbrow,'User Accounts',C['text'],'h2').pack(side='left')
        blue_btn(tbrow,'＋ Add User',lambda: self._add_user_dlg(),w=130,h=38).pack(side='right',padx=4)
        gold_btn(tbrow,'⬆ Bulk Upload (Excel)',lambda: self._bulk_users(),w=190,h=38).pack(side='right',padx=4)

        # Search
        srow=ctk.CTkFrame(root,fg_color=C['bg']); srow.pack(fill='x',padx=24,pady=(0,6))
        sv=ctk.StringVar()
        ctk.CTkEntry(srow,textvariable=sv,placeholder_text='🔎  Search name, email, employee #, department...',
                     width=420,height=36,fg_color=C['surface2'],border_color=C['border'],
                     text_color=C['text'],placeholder_text_color=C['text4'],
                     font=F['body'],corner_radius=6).pack(side='left')
        L(srow,'Usernames are email addresses. Passwords are never shown.',C['text4'],'tiny').pack(side='left',padx=14)

        # Action bar — MUST come before tcard so expand=True doesn't hide it
        abar=ctk.CTkFrame(root,fg_color=C['bg'],height=52); abar.pack(fill='x',padx=24,pady=(0,4)); abar.pack_propagate(False)
        blue_btn(abar,'✎ Edit User',lambda: self._users_edit_selected(),w=130,h=36).pack(side='left',padx=3)
        ghost_btn(abar,'🔑 Reset Password',lambda: self._users_reset_selected(),color=C['gold'],w=160,h=36).pack(side='left',padx=3)
        red_btn(abar,'🗑 Delete',lambda: self._users_delete_selected(),w=110,h=36).pack(side='left',padx=3)
        L(abar,'Tip: click a row to select, then Edit, Reset or Delete. Default admin/staff are protected.',C['text4'],'tiny').pack(side='left',padx=14)

        # Table
        tcard=ctk.CTkFrame(root,fg_color=C['card'],corner_radius=10,border_width=1,border_color=C['border'])
        tcard.pack(fill='both',expand=True,padx=24,pady=(0,8))
        cols=('Employee Number','Employee Name','Email / Username','Department','Role','Status','Availability')
        widths=[140,170,240,140,90,130,110]
        cbt=CheckboxTree(tcard,cols,widths,height=18,simple_mode=True)
        cbt.frame.pack(fill='both',expand=True,padx=2,pady=2)
        self._users_tree=cbt

        # online = logged in within last 5 minutes (LastLogin)
        def load():
            cbt.clear()
            q=sv.get().strip().lower()
            users=qy("""SELECT ISNULL(EmployeeNumber,''), ISNULL(FullName,''), Username,
                               ISNULL(Email,''), ISNULL(Department,''), Role,
                               ISNULL(MustChangePassword,0), LastLogin
                        FROM Users ORDER BY Role, FullName""")
            import datetime
            now=datetime.datetime.now()
            for empno,fn,un,email,dept,role,must,lastlogin in users:
                status='Awaiting first login' if must else 'Active'
                online=False
                if lastlogin:
                    try:
                        online=(now-lastlogin).total_seconds()<300
                    except: online=False
                avail='● Online' if online else '○ Offline'
                vals=(empno or '—', fn or '—', email or un, dept or '—',
                      role.upper(), status, avail)
                if q:
                    hay=' '.join(str(v).lower() for v in vals)
                    if q not in hay: continue
                tag='' if must==0 else 'status_amber'
                cbt.insert(un, vals, tag)
        sv.trace_add('write', lambda *_: load())
        load()

    def _users_reset_selected(self):
        cbt=getattr(self,'_users_tree',None)
        if not cbt: return
        ids=list(cbt.checked) or ([cbt._id_map.get(s) for s in cbt.tree.selection()] if cbt.tree.selection() else [])
        ids=[i for i in ids if i]
        if not ids:
            messagebox.showinfo('Select','Click a user row (or check the box) first.',parent=self); return
        un=ids[0]
        email=(qy("SELECT ISNULL(Email,'') FROM Users WHERE Username=?",(un,)) or [('',)])[0][0]
        self._reset_pw(un,email)

    def _users_delete_selected(self):
        cbt=getattr(self,'_users_tree',None)
        if not cbt: return
        ids=list(cbt.checked) or ([cbt._id_map.get(s) for s in cbt.tree.selection()] if cbt.tree.selection() else [])
        ids=[i for i in ids if i]
        if not ids:
            messagebox.showinfo('Select','Click a user row (or check the box) first.',parent=self); return
        un=ids[0]
        self._delete_user(un)

    def _users_edit_selected(self):
        cbt=getattr(self,'_users_tree',None)
        if not cbt: return
        ids=list(cbt.checked) or ([cbt._id_map.get(s) for s in cbt.tree.selection()] if cbt.tree.selection() else [])
        ids=[i for i in ids if i]
        if not ids:
            messagebox.showinfo('Select','Click a user row (or check the box) first.',parent=self); return
        self._edit_user_dlg(ids[0])

    def _edit_user_dlg(self, username):
        C = T()
        row=qy("""SELECT ISNULL(FullName,''), Username, ISNULL(Email,''), ISNULL(Department,''),
                         Role, ISNULL(EmployeeNumber,''), ISNULL(MustChangePassword,0)
                  FROM Users WHERE Username=?""",(username,))
        if not row:
            messagebox.showerror('Not found','That user no longer exists.',parent=self); self._user_admin(); return
        fn,un,email,dept,role,empno,must = row[0]

        w=ctk.CTkToplevel(self); w.title('Edit User')
        w.geometry('480x620'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,'Edit User',C['text'],'h2').pack(pady=(20,2))
        L(w,un,C['gold'],'sm').pack(pady=(0,10))
        fm=ctk.CTkFrame(w,fg_color=C['surface']); fm.pack(fill='x',padx=40)

        caps_lbl(fm,'Employee Number').pack(anchor='w',pady=(6,3))
        emp_e=inp(fm,w=400); emp_e.insert(0,empno); emp_e.pack()
        caps_lbl(fm,'Full Name').pack(anchor='w',pady=(10,3))
        ne=inp(fm,w=400); ne.insert(0,fn); ne.pack()
        caps_lbl(fm,'Email (username)').pack(anchor='w',pady=(10,3))
        ee=inp(fm,w=400); ee.insert(0,email or un); ee.pack()
        protected = un in ('admin','staff')
        if protected:
            ee.configure(state='disabled')  # don't let the default logins change their username
        caps_lbl(fm,'Department').pack(anchor='w',pady=(10,3))
        dv=ctk.StringVar(value=dept or 'General')
        dept_vals=DEPARTMENTS if (dept in DEPARTMENTS or not dept) else [dept]+DEPARTMENTS
        ctk.CTkComboBox(fm,values=dept_vals,variable=dv,width=400,height=40,
                       fg_color=C['surface2'],border_color=C['border'],button_color=C['gold'],
                       text_color=C['text'],font=F['body'],corner_radius=6,
                       dropdown_fg_color=C['surface2'],dropdown_text_color=C['text']).pack()
        caps_lbl(fm,'Role').pack(anchor='w',pady=(10,3))
        rv=ctk.StringVar(value=role)
        ctk.CTkComboBox(fm,values=['staff','procurement','it','operation','admin','hr','fin'],variable=rv,width=400,height=40,
                       fg_color=C['surface2'],border_color=C['border'],button_color=C['gold'],
                       text_color=C['text'],font=F['body'],corner_radius=6,
                       dropdown_fg_color=C['surface2'],dropdown_text_color=C['text']).pack()

        def save():
            new_name=ne.get().strip(); new_email=ee.get().strip().lower()
            new_dept=dv.get().strip(); new_role=rv.get(); new_emp=emp_e.get().strip()
            if not new_email or '@' not in new_email:
                messagebox.showerror('Invalid','Email/username must be a valid email address.',parent=w); return
            # uniqueness check if email changed
            if new_email!=(email or un).lower():
                if qy("SELECT 1 FROM Users WHERE (Username=? OR Email=?) AND Username<>?",(new_email,new_email,un)):
                    messagebox.showerror('Exists','Another user already has that email.',parent=w); return
            if protected:
                # keep the original username for default accounts
                ex("""UPDATE Users SET FullName=?, Department=?, Role=?, EmployeeNumber=? WHERE Username=?""",
                   (new_name,new_dept,new_role,new_emp,un))
            else:
                ex("""UPDATE Users SET FullName=?, Username=?, Email=?, Department=?, Role=?, EmployeeNumber=? WHERE Username=?""",
                   (new_name,new_email,new_email,new_dept,new_role,new_emp,un))
            w.destroy(); self._user_admin()
        blue_btn(w,'Save Changes',save,w=400,h=44).pack(pady=18,padx=40)

    def _gen_temp_pw(self):
        import random, string
        return 'BSB' + ''.join(random.choices(string.ascii_uppercase+string.digits, k=5))

    def _email_result_popup(self, ok, info, title='Email'):
        C = T()
        w=ctk.CTkToplevel(self); w.title(title)
        w.geometry('520x420'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        if ok is True:
            L(w,'✓ Saved & Email Sent',C['in_fg'],'h2').pack(pady=(18,6))
            L(w,info,C['text3'],'sm').pack(padx=20)
        elif ok is False:
            L(w,'✓ Saved (email failed)',C['low_fg'],'h2').pack(pady=(18,6))
            L(w,info,C['text3'],'sm',wraplength=460,justify='left').pack(padx=20)
        else:
            L(w,'✓ Saved — Email Preview',C['text'],'h2').pack(pady=(18,4))
            L(w,'SMTP is not configured yet, so here is the email to send manually:',C['text3'],'sm').pack(padx=20,pady=(0,8))
            box=ctk.CTkTextbox(w,width=460,height=240,fg_color=C['surface2'],border_color=C['border'],
                              text_color=C['text'],font=F['sm'],corner_radius=6,border_width=1)
            box.pack(padx=20); box.insert('1.0', info); box.configure(state='disabled')
        blue_btn(w,'Done',w.destroy,w=200,h=40).pack(pady=16)

    def _add_user_dlg(self):
        C = T()
        w=ctk.CTkToplevel(self); w.title('Add User')
        w.geometry('480x560'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,'Add New User',C['text'],'h2').pack(pady=(20,4))
        L(w,'A temporary password is generated & emailed to them.',C['text3'],'sm').pack(pady=(0,12))
        fm=ctk.CTkFrame(w,fg_color=C['surface']); fm.pack(fill='x',padx=40)
        caps_lbl(fm,'Full Name').pack(anchor='w',pady=(8,3)); ne=inp(fm,w=400); ne.pack()
        caps_lbl(fm,'Email (this is their username)').pack(anchor='w',pady=(12,3))
        ee=inp(fm,w=400,placeholder=f'name@{EMAIL_DOMAIN}'); ee.pack()
        caps_lbl(fm,'Department').pack(anchor='w',pady=(12,3))
        dv=ctk.StringVar(value='General')
        ctk.CTkComboBox(fm,values=DEPARTMENTS,variable=dv,width=400,height=40,
                       fg_color=C['surface2'],border_color=C['border'],button_color=C['gold'],
                       text_color=C['text'],font=F['body'],corner_radius=6,
                       dropdown_fg_color=C['surface2'],dropdown_text_color=C['text']).pack()
        caps_lbl(fm,'Role').pack(anchor='w',pady=(12,3))
        rv=ctk.StringVar(value='staff')
        ctk.CTkComboBox(fm,values=['staff','procurement','it','operation','admin','hr','fin'],variable=rv,width=400,height=40,
                       fg_color=C['surface2'],border_color=C['border'],button_color=C['gold'],
                       text_color=C['text'],font=F['body'],corner_radius=6,
                       dropdown_fg_color=C['surface2'],dropdown_text_color=C['text']).pack()
        def save():
            name=ne.get().strip(); email=ee.get().strip().lower(); role=rv.get(); dept=dv.get().strip()
            if not email or '@' not in email:
                messagebox.showerror('Invalid','Username must be a valid email address.',parent=w); return
            if qy("SELECT 1 FROM Users WHERE Username=? OR Email=?",(email,email)):
                messagebox.showerror('Exists','A user with that email already exists.',parent=w); return
            # auto employee number: EMP + next sequence
            cnt=(qy("SELECT COUNT(*) FROM Users") or [(0,)])[0][0]
            empno='EMP'+str(cnt+1).zfill(4)
            temp=self._gen_temp_pw()
            ex("""INSERT INTO Users(Username,PasswordHash,Role,FullName,Email,Department,EmployeeNumber,MustChangePassword)
                  VALUES(?,?,?,?,?,?,?,1)""",
               (email,temp,role,name,email,dept,empno))
            ok,info=send_access_email(email,email,temp,parent=w)
            w.destroy(); self._email_result_popup(ok,info,'User Created'); self._user_admin()
        blue_btn(w,'Save User',save,w=400,h=44).pack(pady=18,padx=40)

    def _delete_user(self, username):
        C = T()
        if username in ('admin','staff'):
            messagebox.showinfo('Protected','The default admin/staff accounts cannot be deleted.',parent=self); return
        if messagebox.askyesno('Delete User',f'Delete account "{username}"? This cannot be undone.',parent=self):
            ex("DELETE FROM Users WHERE Username=?",(username,)); self._user_admin()

    def _reset_pw(self, username, email):
        if messagebox.askyesno('Reset Password',
                f'Generate a new temporary password for {username} and email it to them?',parent=self):
            temp=self._gen_temp_pw()
            ex("UPDATE Users SET PasswordHash=?, MustChangePassword=1 WHERE Username=?",(temp,username))
            ok,info=send_access_email(email or username, username, temp, parent=self, reset=True)
            self._email_result_popup(ok,info,'Password Reset'); self._user_admin()

    def _bulk_users(self):
        C = T()
        path=filedialog.askopenfilename(title='Select Excel file',
                filetypes=[('Excel','*.xlsx *.xls')],parent=self)
        if not path: return
        try:
            import openpyxl
        except:
            messagebox.showerror('Missing','Run: pip install openpyxl',parent=self); return
        try:
            wb=openpyxl.load_workbook(path, data_only=True); ws=wb.active
            rows=list(ws.iter_rows(values_only=True))
        except Exception as e:
            messagebox.showerror('Error',f'Could not read file: {e}',parent=self); return

        # Expect columns: Name | Email | Role  (header row optional)
        created=[]; skipped=[]; results=[]
        for r in rows:
            if not r or all(v is None for v in r): continue
            cells=[str(c).strip() if c is not None else '' for c in r]
            # skip header
            if cells[0].lower() in ('name','full name') or (len(cells)>1 and cells[1].lower()=='email'):
                continue
            name=cells[0] if len(cells)>0 else ''
            email=(cells[1] if len(cells)>1 else '').lower()
            role=(cells[2] if len(cells)>2 else 'staff').lower()
            if role not in ALL_ROLES: role='staff'
            if not email or '@' not in email:
                skipped.append(name or '(blank)'); continue
            if qy("SELECT 1 FROM Users WHERE Username=? OR Email=?",(email,email)):
                skipped.append(email); continue
            temp=self._gen_temp_pw()
            ex("INSERT INTO Users(Username,PasswordHash,Role,FullName,Email,MustChangePassword) VALUES(?,?,?,?,?,1)",
               (email,temp,role,name,email))
            ok,info=send_access_email(email,email,temp)
            created.append(email)
            if ok is None: results.append(info)
        # Summary
        msg=f"Created {len(created)} user(s)."
        if skipped: msg+=f"\nSkipped {len(skipped)} (already exist / no email)."
        if results:
            # show first preview block if SMTP off
            self._email_result_popup(None, "\n\n──────────\n\n".join(results[:5]) +
                                     (f"\n\n(+{len(results)-5} more)" if len(results)>5 else ""),
                                     'Bulk Users Created')
        else:
            messagebox.showinfo('Bulk Upload', msg, parent=self)
        self._user_admin()

    # ══ PASSWORD CHANGE (forced first login + anytime) ═══════
    def _force_change_pw(self):
        # Forced screen on first login — cannot skip
        self._change_pw_screen(forced=True)

    def _change_pw_mode(self):
        self._change_pw_screen(forced=False)

    def _change_pw_screen(self, forced):
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        if not forced:
            self._make_settings_bar(self, lambda: self._reload(self._change_pw_mode))
        center=ctk.CTkFrame(self,fg_color=C['bg']); center.place(relx=0.5,rely=0.5,anchor='center')
        card=ctk.CTkFrame(center,fg_color=C['card'],corner_radius=16,border_width=1,border_color=C['border'])
        card.pack()
        inner=ctk.CTkFrame(card,fg_color=C['card']); inner.pack(padx=48,pady=40)
        if forced:
            L(inner,'Set Your Password',C['text'],'h1').pack()
            L(inner,'For security, please choose your own password before continuing.',C['text3'],'sm').pack(pady=(4,24))
        else:
            L(inner,'Change Password',C['text'],'h1').pack()
            L(inner,'Update your account password.',C['text3'],'sm').pack(pady=(4,24))
        caps_lbl(inner,'New Password').pack(anchor='w',pady=(0,4))
        p1=inp(inner,w=360,show='●'); p1.pack(pady=(0,14))
        caps_lbl(inner,'Confirm New Password').pack(anchor='w',pady=(0,4))
        p2=inp(inner,w=360,show='●'); p2.pack(pady=(0,8))
        err=ctk.CTkLabel(inner,text='',text_color=C['red'],font=F['sm']); err.pack(pady=(2,10))
        def save():
            a=p1.get().strip(); b=p2.get().strip()
            if len(a)<4: err.configure(text='Password must be at least 4 characters.'); return
            if a!=b: err.configure(text='Passwords do not match.'); return
            # Update DIRECTLY in SQL
            ex("UPDATE Users SET PasswordHash=?, MustChangePassword=0 WHERE Username=?",(a,self.current_username))
            messagebox.showinfo('Done','Your password has been updated.',parent=self)
            self._modes()
        blue_btn(inner,'Save Password',save,w=360,h=44).pack()
        if not forced:
            ghost_btn(inner,'⬅  Back to Modes',self._modes,w=360,h=38).pack(pady=(12,0))
        else:
            ghost_btn(inner,'Sign Out',self._login,w=360,h=34).pack(pady=(12,0))


    # ══ UNDER CONSTRUCTION ═══════════════════════════════════
    def _under_construction(self):
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        self._make_settings_bar(self, lambda: self._reload(self._under_construction))
        center = ctk.CTkFrame(self, fg_color=C['bg'])
        center.place(relx=0.5, rely=0.5, anchor='center')
        L(center, '🚧', fk='disp').pack(pady=(0,16))
        L(center, 'Under Construction', C['text'], 'h1').pack()
        L(center, 'This feature is coming soon.', C['text3'], 'body').pack(pady=(8,28))
        ghost_btn(center, '⬅  Back', self._modes, w=160).pack()

    def _under_construction_in(self, parent):
        """Render the Under Construction placeholder inside a given frame
        (used by sidebar-based screens where we don't clear the whole window)."""
        C = T()
        for w in parent.winfo_children(): w.destroy()
        center = ctk.CTkFrame(parent, fg_color=C['bg'])
        center.place(relx=0.5, rely=0.5, anchor='center')
        L(center, '🚧', fk='disp').pack(pady=(0,16))
        L(center, 'Under Construction', C['text'], 'h1').pack()
        L(center, 'This feature is coming soon.', C['text3'], 'body').pack(pady=(8,0))

    # ══ ADMIN: DEPARTMENT VIEWS ═══════════════════════════════
    def _dept_screen(self, title, modes_list):
        """Generic department screen used by admin clicking a department card."""
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        self._make_settings_bar(self, lambda: self._reload(lambda: self._dept_screen(title, modes_list)))

        center = ctk.CTkFrame(self, fg_color=C['bg'])
        center.place(relx=0.5, rely=0.5, anchor='center')
        L(center, title, C['text'], 'disp').pack(pady=(0,4))
        L(center, 'ADMIN VIEW', C['gold'], 'caps').pack(pady=(0,36))

        grid = ctk.CTkFrame(center, fg_color=C['bg']); grid.pack()
        cur_row = None
        for idx, (icon, name, desc, cmd) in enumerate(modes_list):
            if idx % 3 == 0:
                cur_row = ctk.CTkFrame(grid, fg_color=C['bg']); cur_row.pack(pady=8)
            card = ctk.CTkFrame(cur_row, fg_color=C['surface'], corner_radius=10,
                               border_width=1, border_color=C['border'], width=200, height=190)
            card.pack(side='left', padx=12); card.pack_propagate(False)
            L(card, icon, fk='disp').pack(pady=(22,6))
            L(card, name, C['text'], 'h3').pack()
            L(card, desc, C['text3'], 'sm').pack(pady=(6,0), padx=14)
            if name == 'Orders':
                pend = (qy("SELECT COUNT(*) FROM Orders WHERE Status='Pending'") or [(0,)])[0][0]
                if pend:
                    ctk.CTkLabel(card, text=f' {pend} new ', text_color='white',
                                 fg_color=C['red'], font=F['caps'], corner_radius=8).pack(pady=(6,0))
            for w in [card]+list(card.winfo_children()):
                w.bind('<Button-1>', lambda e, c2=cmd: c2())
                w.bind('<Enter>', lambda e, w=card: w.configure(border_color=C['gold']))
                w.bind('<Leave>', lambda e, w=card: w.configure(border_color=C['border']))

        hline(center, 24)
        ghost_btn(center, '⬅  Departments', self._modes, w=180).pack()

    def _dept_procurement(self):
        self._dept_screen('Procurement', [
            ('📦','Manage','Add, edit & set quantity', self._manage),
            ('🛒','Orders','See & arrange staff orders', self._orders_mail),
            ('📊','Report','Generate & export reports', self._report),
        ])

    def _dept_it(self):
        self._dept_screen('IT', [
        ('🔧','Manage','Manage IT assets & inventory', self._it_manage),
        ('📤','Assign','Assign devices to users', self._it_assign),
        ('🗑','Clearance','Equipment clearance & disposal', self._under_construction),
    ])

    def _it_assign(self):
        """Device-issuing system: assign IT assets to existing user accounts,
        track who has what, and mark returns."""
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root = ctk.CTkFrame(self, fg_color=C['bg']); root.pack(fill='both', expand=True)
        self._topbar(root, 'IT Department', 'Assign Devices', lambda: self._reload(self._it_assign))

        fbar = ctk.CTkFrame(root, fg_color=C['topbar'], corner_radius=0, height=60)
        fbar.pack(fill='x'); fbar.pack_propagate(False)
        hline(fbar, 0)
        af = {'status':'Assigned'}
        sv = ctk.StringVar()
        L(fbar,'Search',C['text4'],'xs').pack(side='left',padx=(18,8),pady=18)
        ctk.CTkEntry(fbar, textvariable=sv, placeholder_text='Search person or device...',
                    width=240, height=36, fg_color=C['surface2'], border_color=C['border'],
                    text_color=C['text'], placeholder_text_color=C['text4'],
                    font=F['body'], corner_radius=6).pack(side='left',pady=12)
        blue_btn(fbar,'＋ Assign a Device', lambda: self._it_assign_dlg(load), w=170, h=38).pack(side='right',padx=14,pady=11)

        chips = ctk.CTkFrame(fbar, fg_color=C['topbar']); chips.pack(side='left', padx=20)
        chip_btns = {}
        def set_status(s):
            af['status']=s; load()
            for k,b in chip_btns.items():
                on=(k==s)
                b.configure(fg_color=C['gold'] if on else C['surface2'],
                            text_color='#0B1F3A' if on else C['text3'])
        for label,key in [('Currently Assigned','Assigned'),('Returned','Returned'),('All','all')]:
            b=ctk.CTkButton(chips,text=label,height=30,width=130,corner_radius=15,
                            fg_color=C['gold'] if key=='Assigned' else C['surface2'],
                            text_color='#0B1F3A' if key=='Assigned' else C['text3'],
                            hover_color=C['gold2'], font=F['xs'],
                            command=lambda k=key:set_status(k))
            b.pack(side='left',padx=4); chip_btns[key]=b

        tcard = ctk.CTkFrame(root, fg_color=C['card'], corner_radius=12, border_width=1, border_color=C['border'])
        tcard.pack(fill='both', expand=True, padx=14, pady=12)
        cbt = CheckboxTree(tcard,
            ('Device','Category','Assigned To','Username','Assigned By','Date','Due','Status'),
            [200,130,150,110,120,90,90,90], height=18, simple_mode=True)
        cbt.frame.pack(fill='both', expand=True, padx=2, pady=2)
        self._assign_tree = cbt

        bbar = ctk.CTkFrame(root, fg_color=C['bg'], height=52); bbar.pack(fill='x', padx=14, pady=(0,10)); bbar.pack_propagate(False)
        gold_btn(bbar,'↩ Mark Returned', lambda: self._it_return(load), w=150, h=38).pack(side='left', padx=4)
        L(bbar,'Check assignment(s) (☐ in the header selects all), then mark them returned.',C['text4'],'tiny').pack(side='left',padx=14)

        def load():
            cbt.clear()
            srch=f"%{sv.get()}%"
            base="""SELECT AssignID, ISNULL(AssetLabel,''), ISNULL(CatName,''), ISNULL(AssignedTo,''),
                           ISNULL(AssignedToUsername,''), ISNULL(AssignedBy,''),
                           CONVERT(varchar(10),AssignDate,23), ISNULL(DueDate,''), ISNULL(Status,'')
                    FROM (
                      SELECT a.*, c.Name AS CatName
                      FROM ITAssignments a
                      LEFT JOIN ITAssets ast ON a.AssetID=ast.AssetID
                      LEFT JOIN ITCategories c ON ast.ITCatID=c.ITCatID
                    ) x
                    WHERE (AssetLabel LIKE ? OR AssignedTo LIKE ? OR AssignedToUsername LIKE ?) """
            if af['status']!='all':
                base+="AND Status=? ORDER BY AssignID DESC"
                rows=qy(base,(srch,srch,srch,af['status']))
            else:
                base+="ORDER BY AssignID DESC"
                rows=qy(base,(srch,srch,srch))
            for r in rows:
                cbt.insert(r[0], r[1:])
        sv.trace_add('write', lambda *_: load())
        load()

    def _it_assign_dlg(self, reload_fn):
        """Full assign flow:
        1. Assign Type (Permanent / Temporary) — compulsory, unlocks the rest
        2. Category (VR, Chromebooks, ...) — filters devices
        3. Device (from that category)
        4. Department
        5. User (by email) → fills Employee Number automatically
        6. Due Date (mandatory if Temporary, optional if Permanent)
        7. Notes (optional) — if filled, emailed to the assignee
        An email is sent to the assignee on assign."""
        C = T()
        w = ctk.CTkToplevel(self); w.title('Assign a Device')
        w.geometry('560x760'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,'Assign a Device',C['text'],'h2').pack(pady=(16,2))
        L(w,'Choose an assign type to begin',C['text4'],'sm').pack(pady=(0,6))

        fm = ctk.CTkScrollableFrame(w,fg_color=C['surface'],width=500,height=560)
        fm.pack(fill='both',expand=True,padx=28,pady=(4,0))

        WIDE = 460
        # ── 1. Assign Type ──
        caps_lbl(fm,'① Assign Type  (required)').pack(anchor='w',pady=(6,3))
        type_var = ctk.StringVar(value='')
        trow = ctk.CTkFrame(fm, fg_color=C['surface']); trow.pack(anchor='w', fill='x')
        # ── widgets that start disabled ──
        cat_box = dev_box = dept_box = user_box = None
        emp_entry = due_entry = notes_entry = None

        # data sources
        cats = qy("SELECT ITCatID, Name FROM ITCategories ORDER BY Name")
        cat_opts = [r[1] for r in cats]
        cat_id_by_name = {r[1]: r[0] for r in cats}
        users = qy("""SELECT UserID, ISNULL(FullName,''), Username, ISNULL(Email,''),
                             ISNULL(EmployeeNumber,''), ISNULL(Department,'')
                      FROM Users ORDER BY FullName""")
        # user identified by email (fallback username)
        user_opts = [ (r[3] or r[2]) for r in users ]
        user_by_key = { (r[3] or r[2]): r for r in users }
        dept_opts = sorted({ (r[5] or '').strip() for r in users if (r[5] or '').strip() }) or ['General']

        def set_state(widget, state):
            try: widget.configure(state=state)
            except: pass

        def refresh_devices(*_):
            cn = cat_var.get()
            cid = cat_id_by_name.get(cn)
            if cid is None:
                dev_box.configure(values=[]); dev_var.set(''); return
            devs = qy("""SELECT a.AssetID, ISNULL(a.AssetTag,''), ISNULL(a.Name,''), ISNULL(a.Model,'')
                         FROM ITAssets a WHERE a.ITCatID=? ORDER BY a.AssetID""",(cid,))
            opts=[]; dmap={}
            for aid,tag,nm,model in devs:
                label = f"{tag or nm or model or ('#'+str(aid))}"
                if model and model not in label: label += f"  ·  {model}"
                label += f"  (#{aid})"
                opts.append(label); dmap[label]=(aid, f"{tag or nm} ({cn})")
            dev_box.configure(values=opts)
            dev_var.set(opts[0] if opts else '')
            refresh_devices.map = dmap
        refresh_devices.map = {}

        def on_user_pick(*_):
            key = user_var.get()
            row = user_by_key.get(key)
            if row and emp_entry:
                emp_entry.configure(state='normal')
                emp_entry.delete(0,'end')
                emp_entry.insert(0, row[4] or '—')
                emp_entry.configure(state='disabled')

        def enable_all():
            for wdg in [cat_box, dev_box, dept_box, user_box, due_entry, notes_entry]:
                set_state(wdg, 'normal')
            refresh_devices()
            on_user_pick()
            # due date label reflects mandatory/optional
            due_lbl.configure(text=('④ Due Date  (REQUIRED for temporary)'.upper()
                                    if type_var.get()=='Temporary'
                                    else '④ Due Date  (optional)'.upper()))

        def pick_type(t):
            type_var.set(t)
            for b,bt in type_btns:
                on=(bt==t)
                b.configure(fg_color=C['gold'] if on else C['surface2'],
                            text_color='#0B1F3A' if on else C['text3'])
            enable_all()

        type_btns=[]
        for t in ['Permanent','Temporary']:
            b=ctk.CTkButton(trow,text=t,width=210,height=42,corner_radius=8,
                            fg_color=C['surface2'], text_color=C['text3'],
                            hover_color=C['border'], font=F['sm_b'],
                            command=lambda tt=t: pick_type(tt))
            b.pack(side='left',padx=(0,10),pady=4); type_btns.append((b,t))

        # ── 2. Category ──
        caps_lbl(fm,'② Category').pack(anchor='w',pady=(12,3))
        cat_var = ctk.StringVar(value=cat_opts[0] if cat_opts else '')
        cat_box = ctk.CTkComboBox(fm, values=cat_opts, variable=cat_var, width=WIDE, height=40,
                       fg_color=C['surface2'], border_color=C['border'], button_color=C['gold'],
                       button_hover_color=C['gold2'], text_color=C['text'], font=F['body'],
                       corner_radius=6, dropdown_fg_color=C['surface2'], dropdown_text_color=C['text'],
                       command=refresh_devices, state='disabled')
        cat_box.pack()

        # ── 3. Device ──
        caps_lbl(fm,'③ Device').pack(anchor='w',pady=(12,3))
        dev_var = ctk.StringVar(value='')
        dev_box = ctk.CTkComboBox(fm, values=[], variable=dev_var, width=WIDE, height=40,
                       fg_color=C['surface2'], border_color=C['border'], button_color=C['gold'],
                       button_hover_color=C['gold2'], text_color=C['text'], font=F['body'],
                       corner_radius=6, dropdown_fg_color=C['surface2'], dropdown_text_color=C['text'],
                       state='disabled')
        dev_box.pack()

        # ── Department ──
        caps_lbl(fm,'Department').pack(anchor='w',pady=(12,3))
        dept_var = ctk.StringVar(value=dept_opts[0] if dept_opts else 'General')
        dept_box = ctk.CTkComboBox(fm, values=dept_opts, variable=dept_var, width=WIDE, height=40,
                       fg_color=C['surface2'], border_color=C['border'], button_color=C['gold'],
                       button_hover_color=C['gold2'], text_color=C['text'], font=F['body'],
                       corner_radius=6, dropdown_fg_color=C['surface2'], dropdown_text_color=C['text'],
                       state='disabled')
        dept_box.pack()

        # ── User (by email) ──
        caps_lbl(fm,'Select User  (email)').pack(anchor='w',pady=(12,3))
        user_var = ctk.StringVar(value=user_opts[0] if user_opts else '')
        user_box = ctk.CTkComboBox(fm, values=user_opts, variable=user_var, width=WIDE, height=40,
                       fg_color=C['surface2'], border_color=C['border'], button_color=C['gold'],
                       button_hover_color=C['gold2'], text_color=C['text'], font=F['body'],
                       corner_radius=6, dropdown_fg_color=C['surface2'], dropdown_text_color=C['text'],
                       command=on_user_pick, state='disabled')
        user_box.pack()

        # ── Employee Number (auto-filled, read-only) ──
        caps_lbl(fm,"Employee Number  (auto-filled)").pack(anchor='w',pady=(12,3))
        emp_entry = inp(fm, w=WIDE); emp_entry.pack(); emp_entry.configure(state='disabled')

        # ── Due Date ──
        due_lbl = caps_lbl(fm,'④ Due Date  (optional)'); due_lbl.pack(anchor='w',pady=(12,3))
        due_entry = inp(fm, w=WIDE, placeholder='YYYY-MM-DD'); due_entry.pack()
        due_entry.configure(state='disabled')

        # ── Notes ──
        caps_lbl(fm,'Notes  (optional — emailed if filled)').pack(anchor='w',pady=(12,3))
        notes_entry = inp(fm, w=WIDE); notes_entry.pack()
        notes_entry.configure(state='disabled')

        def save():
            try:
                if not type_var.get():
                    raise ValueError('Pick an assign type (Permanent or Temporary) first.')
                dmap = refresh_devices.map
                if dev_var.get() not in dmap:
                    raise ValueError('Pick a device.')
                aid, alabel = dmap[dev_var.get()]
                urow = user_by_key.get(user_var.get())
                if not urow:
                    raise ValueError('Pick a user.')
                uid, ufull, uname, uemail, empno, _udept = urow
                assign_type = type_var.get()
                due = due_entry.get().strip()
                if assign_type=='Temporary' and not due:
                    raise ValueError('Due date is required for a Temporary assignment.')
                notes = notes_entry.get().strip()
                dept = dept_var.get().strip()
                cat = cat_var.get()
                ex("""INSERT INTO ITAssignments(AssetID,AssetLabel,UserID,AssignedTo,AssignedToUsername,
                       AssignedToEmail,AssignedBy,DueDate,Status,Notes,AssignType,CategoryName,EmployeeNumber)
                       VALUES(?,?,?,?,?,?,?,?,'Assigned',?,?,?,?)""",
                   (aid, alabel, uid, ufull, uname, uemail, self.current_user, due,
                    notes, assign_type, cat, empno))
                # send email
                self._send_assign_email(uemail, ufull, alabel, assign_type, due, notes)
                reload_fn(); w.destroy()
                messagebox.showinfo('Assigned',
                    f'{alabel} assigned to {ufull}.\nAn email notification was sent.', parent=self)
            except Exception as e:
                messagebox.showerror('Error', str(e), parent=w)
        blue_btn(w,'Assign Device',save,w=460,h=44).pack(pady=14,padx=40)

    def _send_assign_email(self, to_email, full_name, device_label, assign_type, due, notes):
        """Notify a user that a device has been assigned to them."""
        if not to_email or not SMTP_CONFIG.get('enabled'):
            return
        try:
            import smtplib
            from email.mime.text import MIMEText
            lines = [f"Dear {full_name},", "",
                     "A device has been assigned to you:", "",
                     f"  Device: {device_label}",
                     f"  Assignment: {assign_type}"]
            if due: lines.append(f"  Return by: {due}")
            if notes: lines += ["", "Notes:", f"  {notes}"]
            lines += ["", "Please collect it from the IT office.", "", "BSB IT Department"]
            msg = MIMEText("\n".join(lines))
            msg['Subject'] = 'BSB IT — A device has been assigned to you'
            msg['From'] = SMTP_CONFIG['sender']; msg['To'] = to_email
            s = smtplib.SMTP(SMTP_CONFIG['host'], SMTP_CONFIG['port'], timeout=15)
            s.starttls(); s.login(SMTP_CONFIG['sender'], SMTP_CONFIG['password'])
            s.sendmail(SMTP_CONFIG['sender'], [to_email], msg.as_string()); s.quit()
        except Exception as e:
            print(f"Assign email error: {e}")

    def _it_return(self, reload_fn):
        cbt = getattr(self, '_assign_tree', None)
        if cbt is None: return
        ids = list(cbt.checked)
        if not ids:
            sel = cbt.tree.selection()
            if sel:
                iid = cbt._id_map.get(sel[0])
                if iid is not None: ids = [iid]
        if not ids:
            messagebox.showinfo('Select','Check the assignment(s) you want to mark returned (or click a row).',parent=self); return
        already = 0; done = 0
        for aid in ids:
            vals = cbt._val_map.get(aid)
            status = vals[7] if vals and len(vals) > 7 else ''
            if str(status) == 'Returned':
                already += 1; continue
            ex("UPDATE ITAssignments SET Status='Returned', ReturnDate=GETDATE() WHERE AssignID=?",(aid,))
            done += 1
        reload_fn()
        if already and not done:
            messagebox.showinfo('Already returned','Selected device(s) were already marked returned.',parent=self)

    def _dept_hr(self):
        self._dept_screen('HR', [
            ('📦','Manage','HR records management', self._under_construction),
            ('📊','Report','Generate & export reports', self._under_construction),
        ])

    def _dept_fin(self):
        self._dept_screen('Finance', [
            ('💳','Payment Plan','Manage payment plans', self._fin_payment_plan),
        ])

    def _fin_payment_plan(self):
        """Finance → Payment Plan screen with a left sidebar for sub-modes."""
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root = ctk.CTkFrame(self, fg_color=C['bg']); root.pack(fill='both', expand=True)
        back_fn = self._dept_fin if self.current_role == 'admin' else self._modes
        self._topbar(root, 'Finance', 'Payment Plan', lambda: self._reload(self._fin_payment_plan))

        # ── Left sidebar with sub-modes ─────────────────────────
        body = ctk.CTkFrame(root, fg_color=C['bg']); body.pack(fill='both', expand=True)

        sb = tk.Frame(body, bg=C['sidebar'], width=230)
        sb.pack(side='left', fill='y'); sb.pack_propagate(False)
        tk.Frame(sb, bg='#1A3A5C', width=1).pack(side='right', fill='y')

        tk.Label(sb, text='PAYMENT PLAN', bg=C['sidebar'], fg=C['side_mut'],
                 font=('Outfit', 9, 'bold'), anchor='w', padx=16, pady=14).pack(fill='x')
        tk.Frame(sb, bg='#132944', height=1).pack(fill='x')

        content = ctk.CTkFrame(body, fg_color=C['bg']); content.pack(side='left', fill='both', expand=True)
        active_btn = [None]

        FIN_SUBMODES = [
            ('💳', 'Payment Plan', 'payment'),
        ]

        def set_active(btn):
            if active_btn[0]:
                active_btn[0].configure(bg=C['sidebar'], fg=C['side_mut'],
                                        activebackground=C['sidebar'])
            btn.configure(bg='#132944', fg=C['gold'],
                          activebackground='#132944')
            active_btn[0] = btn

        def show_submode(key):
            for w in content.winfo_children(): w.destroy()
            # All sub-modes are Under Construction for now
            self._under_construction_in(content)

        # Build sub-mode buttons and track them
        sub_buttons = []
        for icon, label, key in FIN_SUBMODES:
            btn = tk.Button(sb, text=f'  {icon}  {label}',
                           bg=C['sidebar'], fg=C['side_mut'],
                           font=('Outfit', 12, 'bold'), relief='flat',
                           anchor='w', padx=12, pady=10, cursor='hand2',
                           bd=0, highlightthickness=0,
                           activebackground='#132944', activeforeground=C['gold'],
                           wraplength=190, justify='left')
            btn.pack(fill='x')
            btn.configure(command=lambda b=btn, k=key: (set_active(b), show_submode(k)))
            sub_buttons.append((btn, key))

        tk.Frame(sb, bg='#132944', height=1).pack(fill='x', pady=(8,0))
        tk.Button(sb, text='⬅  Back',
                  bg=C['sidebar'], fg=C['side_mut'],
                  font=('Outfit', 11), relief='flat',
                  anchor='w', padx=16, pady=10, cursor='hand2',
                  bd=0, highlightthickness=0,
                  activebackground='#132944', activeforeground=C['text'],
                  command=lambda: self._reload(back_fn)).pack(fill='x', pady=(4,0))

        # Auto-select first sub-mode
        if sub_buttons:
            set_active(sub_buttons[0][0])
            show_submode(sub_buttons[0][1])



    def _op_manage(self):
        """Operations PPE register — its own structure (17 columns) + bulk upload."""
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root = ctk.CTkFrame(self, fg_color=C['bg']); root.pack(fill='both', expand=True)
        self._topbar(root, 'Operations', 'PPE Asset Register', lambda: self._reload(self._op_manage))

        tb = ctk.CTkFrame(root, fg_color=C['topbar'], corner_radius=0, height=60)
        tb.pack(fill='x'); tb.pack_propagate(False); hline(tb,0)
        sv = ctk.StringVar()
        L(tb,'Search',C['text4'],'xs').pack(side='left',padx=(18,8),pady=18)
        ctk.CTkEntry(tb, textvariable=sv, placeholder_text='Search description, code, serial, location...',
                    width=300, height=36, fg_color=C['surface2'], border_color=C['border'],
                    text_color=C['text'], placeholder_text_color=C['text4'],
                    font=F['body'], corner_radius=6).pack(side='left',pady=12)
        blue_btn(tb,'＋ Add Asset',lambda: self._op_asset_dlg(load),w=120,h=36).pack(side='right',padx=4,pady=12)
        gold_btn(tb,'⬆ Bulk Upload',lambda: self._op_bulk(load),w=130,h=36).pack(side='right',padx=4,pady=12)
        ghost_btn(tb,'✎ Edit',lambda: self._op_asset_dlg(load,True),w=80,h=36).pack(side='right',padx=4,pady=12)
        red_btn(tb,'🗑 Delete',lambda: self._op_del(load),w=90,h=36).pack(side='right',padx=4,pady=12)

        hint = ctk.CTkFrame(root, fg_color=C['bg'], height=26); hint.pack(fill='x'); hint.pack_propagate(False)
        L(hint,'ⓘ  Check items to select (☐ in the header selects/deselects all) · Separate from Procurement and IT.',C['text4'],'tiny').pack(side='left',padx=18,pady=4)

        tcard = ctk.CTkFrame(root, fg_color=C['card'], corner_radius=12, border_width=1, border_color=C['border'])
        tcard.pack(fill='both', expand=True, padx=14, pady=(2,14))
        cbt = CheckboxTree(tcard,
            ('Asset Number','Asset Code','Description','Category','Sub Cat','Location','Condition','Serial','Brand','Model'),
            [110,100,200,110,110,130,90,120,100,100], height=18, simple_mode=True)
        cbt.frame.pack(fill='both', expand=True, padx=2, pady=2)
        self._op_tree = cbt

        def load():
            cbt.clear()
            srch=f"%{sv.get()}%"
            rows = qy("""SELECT OpsID, ISNULL(AssetNumber,''), ISNULL(AssetCode,''), ISNULL(AssetDescription,''),
                                ISNULL(Category,''), ISNULL(SubCategory,''), ISNULL(Location,''),
                                ISNULL(Condition,''), ISNULL(SerialCode,''), ISNULL(Brand,''), ISNULL(Model,'')
                         FROM OpsAssets
                         WHERE AssetDescription LIKE ? OR AssetCode LIKE ? OR SerialCode LIKE ?
                            OR Location LIKE ? OR AssetNumber LIKE ?
                         ORDER BY OpsID DESC""",(srch,srch,srch,srch,srch))
            for r in rows:
                cbt.insert(r[0], r[1:])
        cbt.tree.bind('<Double-1>', lambda e: self._op_asset_dlg(load, True))
        sv.trace_add('write', lambda *_: load())
        load()

    OP_FIELDS = [('Asset Number','AssetNumber'),('Asset Code','AssetCode'),
                 ('Asset Description','AssetDescription'),('LPO Number','LPONumber'),
                 ('Country','Country'),('City','City'),('Branch','Branch'),
                 ('Location','Location'),('Sub Location','SubLocation'),
                 ('Category','Category'),('Sub Category','SubCategory'),
                 ('Details','Details'),('Insert Date','InsertDate'),
                 ('Condition','Condition'),('Serial Code','SerialCode'),
                 ('Brand','Brand'),('Model','Model')]

    def _op_tree_selection(self, multi=False):
        """Return checked OpsAsset ID(s), falling back to the highlighted row."""
        cbt = getattr(self, '_op_tree', None)
        if cbt is None: return [] if multi else None
        ids = list(cbt.checked)
        if not ids:
            sel = cbt.tree.selection()
            if sel:
                iid = cbt._id_map.get(sel[0])
                if iid is not None: ids = [iid]
        if multi: return ids
        return ids[0] if ids else None

    def _op_asset_dlg(self, reload_fn, edit=False):
        C = T()
        editing = None; existing = None
        if edit:
            editing = self._op_tree_selection()
            if editing is None: messagebox.showinfo('Select','Tick an asset (or click its row) to edit.',parent=self); return
            cols = ','.join(c for _,c in self.OP_FIELDS)
            r = qy(f"SELECT {cols} FROM OpsAssets WHERE OpsID=?",(editing,))
            if r: existing = r[0]

        w = ctk.CTkToplevel(self); w.title('Edit Asset' if editing else 'Add Asset')
        w.geometry('540x720'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['gold']).pack(fill='x')
        L(w,'Edit Operations Asset' if editing else 'Add Operations Asset',C['text'],'h2').pack(pady=(16,4))
        fm = ctk.CTkScrollableFrame(w,fg_color=C['surface'],width=480,height=560); fm.pack(fill='both',expand=True,padx=28,pady=(6,4))

        widgets = {}
        for i,(label,col) in enumerate(self.OP_FIELDS):
            caps_lbl(fm, label).pack(anchor='w', pady=(9,3))
            e = inp(fm, w=440)
            if existing and existing[i]: e.insert(0, str(existing[i]))
            e.pack(); widgets[col]=e

        def save():
            try:
                cols = [c for _,c in self.OP_FIELDS]
                vals = [widgets[c].get().strip() for c in cols]
                if not vals[2]:  # description required
                    raise ValueError('Asset Description is required')
                if editing:
                    setclause = ','.join(f"{c}=?" for c in cols)
                    ex(f"UPDATE OpsAssets SET {setclause} WHERE OpsID=?", tuple(vals)+(editing,))
                else:
                    placeholders = ','.join('?' for _ in cols)
                    ex(f"INSERT INTO OpsAssets({','.join(cols)}) VALUES({placeholders})", tuple(vals))
                reload_fn(); w.destroy()
            except Exception as e:
                messagebox.showerror('Error',str(e),parent=w)
        blue_btn(w,'Save Asset',save,w=480,h=44).pack(pady=(4,16),padx=30)

    def _op_del(self, reload_fn):
        ids = self._op_tree_selection(multi=True)
        if not ids: messagebox.showinfo('Select','Tick the asset(s) you want to delete (or click a row).',parent=self); return
        C = T()
        w = ctk.CTkToplevel(self); w.title('Confirm Delete')
        w.geometry('380x200'); w.configure(fg_color=C['surface']); w.grab_set(); w.after(60,w.lift)
        ctk.CTkFrame(w,height=3,fg_color=C['red']).pack(fill='x')
        L(w,'Delete Operations assets?',C['text'],'h2').pack(pady=(18,8))
        L(w,f'{len(ids)} asset(s) will be permanently deleted.',C['text3'],'sm').pack()
        bf = ctk.CTkFrame(w,fg_color=C['surface']); bf.pack(pady=18)
        ghost_btn(bf,'Cancel',w.destroy,w=110,h=38).pack(side='left',padx=8)
        def confirm():
            for i in ids: ex("DELETE FROM OpsAssets WHERE OpsID=?",(i,))
            reload_fn(); w.destroy()
        red_btn(bf,'Delete All',confirm,w=120,h=38).pack(side='left',padx=8)

    def _op_bulk(self, reload_fn):
        """Operations bulk upload — matches the PPE Master File layout (17 columns)."""
        path = filedialog.askopenfilename(title='Select Operations PPE Excel',
                filetypes=[('Excel','*.xlsx *.xls')], parent=self)
        if not path: return
        try:
            import openpyxl
        except:
            messagebox.showerror('Missing','Run: pip install openpyxl',parent=self); return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            # prefer the 'Master File' sheet; else first sheet
            ws = wb['Master File'] if 'Master File' in wb.sheetnames else wb.active
            rows = list(ws.iter_rows(values_only=True))
        except Exception as e:
            messagebox.showerror('Error',f'Could not read file: {e}',parent=self); return

        cols = [c for _,c in self.OP_FIELDS]
        labels = [l.lower() for l,_ in self.OP_FIELDS]
        # find header row
        hr = None
        for i,row in enumerate(rows[:8]):
            cells=[str(c).strip().lower() if c is not None else '' for c in row]
            if sum(1 for c in cells if c in labels) >= 5:
                hr=i; break
        if hr is None: hr=0
        header=[str(c).strip().lower() if c is not None else '' for c in rows[hr]]
        # map each db column to a sheet column index by matching label
        col_idx = {}
        for ci,hc in enumerate(header):
            for li,lab in enumerate(labels):
                if hc==lab or (hc and (hc in lab or lab in hc)):
                    col_idx[cols[li]] = ci; break

        added=0; skipped=0
        for row in rows[hr+1:]:
            if not row or all(v is None for v in row): continue
            cells=[str(c).strip() if c is not None else '' for c in row]
            vals=[]
            for c in cols:
                ci = col_idx.get(c)
                vals.append(cells[ci] if (ci is not None and ci < len(cells)) else '')
            if not vals[2]:  # need a description
                skipped+=1; continue
            placeholders=','.join('?' for _ in cols)
            ex(f"INSERT INTO OpsAssets({','.join(cols)}) VALUES({placeholders})", tuple(vals))
            added+=1
        msg=f'Imported {added} Operations asset(s).'
        if skipped: msg+=f'\nSkipped {skipped} row(s) with no description.'
        messagebox.showinfo('Operations Bulk Upload', msg, parent=self)
        reload_fn()

    def _dept_operation(self):
        self._dept_screen('Operation', [
            ('📋','Assign','Assign operational tasks', self._under_construction),
            ('⚙️','Manage','Manage operational assets (PPE register)', self._op_manage),
        ])

    def _dept_staff(self):
        """Admin clicking Staff → goes directly to general staff modes."""
        self._dept_general_staff()

    def _dept_general_staff(self):
        """Admin viewing General Staff modes - same as a staff member sees."""
        self._dept_screen('General Staff', [
            ('🛒','Order','Order stationery for class', self._order),
            ('🕘','Orders History','See order history & status', self._order_history),
            ('🔑','Password','Change password', self._change_pw_mode),
        ])

    def _dept_admin(self):
        """Admin dept - only Staff History."""
        self._dept_screen('Admin', [
            ('📋','Staff History','View all users order\n& issue history', self._staff_history),
        ])

    # ══ ADMIN: STAFF HISTORY ══════════════════════════════════
    def _staff_history(self):
        self.state('zoomed')
        C = T()
        self._clear(); self.configure(fg_color=C['bg'])
        root = ctk.CTkFrame(self, fg_color=C['bg']); root.pack(fill='both', expand=True)
        self._topbar(root, 'Admin', 'Staff History', lambda: self._reload(self._staff_history))
        body = ctk.CTkFrame(root, fg_color=C['bg']); body.pack(fill='both', expand=True)

        # Left: user list
        left = ctk.CTkFrame(body, fg_color=C['surface'], corner_radius=0, width=280)
        left.pack(side='left', fill='y'); left.pack_propagate(False)
        ctk.CTkFrame(left, height=1, fg_color=C['border']).pack(fill='x', side='bottom')
        row_lbl = ctk.CTkFrame(left, fg_color=C['surface']); row_lbl.pack(fill='x', padx=12, pady=(12,0))
        L(row_lbl, 'All Users', C['text'], 'h3').pack(side='left', padx=6)
        blue_btn(row_lbl,'📊 Export',lambda: self._e_staff_history(),w=100,h=30).pack(side='right',padx=4)
        ctk.CTkFrame(left, height=1, fg_color=C['border']).pack(fill='x', pady=(10,0))

        users_lb = tk.Listbox(left, bg=C['surface'], fg=C['text'], font=('Outfit',13,'bold'),
                             selectbackground=C['blue'], selectforeground='white',
                             height=30, bd=0, highlightthickness=0, relief='flat', activestyle='none')
        users_lb.pack(fill='both', expand=True, padx=8, pady=8)

        users = [(r[0],r[1],r[2]) for r in qy("SELECT Username, ISNULL(FullName,''), Role, ISNULL(Email,'') FROM Users ORDER BY Role, FullName")]
        for un, fn, role in users:
            users_lb.insert('end', f"  {fn or un}  ({role})")

        # Right: history panel
        right = ctk.CTkScrollableFrame(body, fg_color=C['bg'])
        right.pack(side='left', fill='both', expand=True, padx=20, pady=16)

        def show_user(event=None):
            # Clear right
            for w in right.winfo_children(): w.destroy()
            sel = users_lb.curselection()
            if not sel: return
            idx = sel[0]
            un, fn, role = users[idx]
            display = fn or un

            L(right, f'{display}', C['text'], 'h2').pack(anchor='w', pady=(0,2))
            L(right, f'{role.upper()}  ·  {un}', C['text4'], 'sm').pack(anchor='w', pady=(0,16))

            # Orders
            L(right, '🛒  Order History', C['text'], 'h3').pack(anchor='w', pady=(0,8))
            orders = qy("""SELECT o.OrderID, o.Department, o.Status, o.CreatedAt
                           FROM Orders o WHERE o.OrderedBy=? ORDER BY o.CreatedAt DESC""", (un,))
            if not orders:
                L(right, 'No orders placed.', C['text3'], 'sm').pack(anchor='w', pady=(0,12))
            for oid, dept, status, created in orders:
                card = ctk.CTkFrame(right, fg_color=C['card'], corner_radius=8,
                                   border_width=1, border_color=C['border'])
                card.pack(fill='x', pady=4)
                hd = ctk.CTkFrame(card, fg_color=C['card']); hd.pack(fill='x', padx=14, pady=(10,2))
                dt = created.strftime('%d/%m/%Y %H:%M') if hasattr(created,'strftime') else str(created)
                L(hd, f'Order #{oid}', C['text'], 'sm_b').pack(side='left')
                sc = {'Pending':C['low_fg'],'Arranged':C['in_fg'],'Collected':C['text3']}.get(status,C['text3'])
                ctk.CTkLabel(hd, text=f' {status} ', text_color=sc,
                            font=F['caps'], corner_radius=6).pack(side='right')
                L(hd, dt, C['text4'], 'tiny').pack(side='right', padx=8)
                items = qy("SELECT ItemName, Quantity FROM OrderItems WHERE OrderID=?", (oid,))
                if dept: L(card, f'Dept: {dept}', C['text3'], 'tiny').pack(anchor='w', padx=14, pady=(2,0))
                irow2 = ctk.CTkFrame(card, fg_color=C['card']); irow2.pack(anchor='w', padx=14, pady=(4,10))
                for nm,q in items:
                    ctk.CTkLabel(irow2,text=f'{q}×',text_color=C['gold'],
                                font=ctk.CTkFont('Outfit',15,'bold')).pack(side='left',padx=(0,4))
                    ctk.CTkLabel(irow2,text=nm,text_color=C['text2'],
                                font=ctk.CTkFont('Outfit',13)).pack(side='left',padx=(0,14))

            ctk.CTkFrame(right, height=1, fg_color=C['border']).pack(fill='x', pady=12)

            # Issues
            L(right, '📋  Issue History', C['text'], 'h3').pack(anchor='w', pady=(0,8))
            issues = qy("""SELECT r.IssueID, r.RecipientName, r.Department, r.IssuedAt, ISNULL(r.Completed,0)
                           FROM IssueRecords r WHERE r.IssuedBy=? ORDER BY r.IssuedAt DESC""", (fn or un,))
            # Also search by username
            if not issues:
                issues = qy("""SELECT r.IssueID, r.RecipientName, r.Department, r.IssuedAt, ISNULL(r.Completed,0)
                               FROM IssueRecords r WHERE r.RecipientName LIKE ?
                               ORDER BY r.IssuedAt DESC""", (f'%{fn}%',))
            if not issues:
                L(right, 'No issues recorded.', C['text3'], 'sm').pack(anchor='w')
            for iid, recip, dept, at, done in issues:
                card = ctk.CTkFrame(right, fg_color=C['card'], corner_radius=8,
                                   border_width=1, border_color=C['border'])
                card.pack(fill='x', pady=4)
                hd = ctk.CTkFrame(card, fg_color=C['card']); hd.pack(fill='x', padx=14, pady=(10,2))
                dt = at.strftime('%d/%m/%Y %H:%M') if hasattr(at,'strftime') else str(at)
                L(hd, f'Issue #{iid}  —  {recip}', C['text'], 'sm_b').pack(side='left')
                stxt = '✓ Done' if done else 'Remaining'
                sc2 = C['in_fg'] if done else C['low_fg']
                ctk.CTkLabel(hd, text=f' {stxt} ', text_color=sc2, font=F['caps'], corner_radius=6).pack(side='right')
                L(hd, dt, C['text4'], 'tiny').pack(side='right', padx=8)
                iitems = qy("SELECT ItemName, Quantity FROM IssueItems WHERE IssueID=?", (iid,))
                itxt2 = '  •  '.join(f"{q}× {n}" for n,q in iitems)
                if dept: itxt2 = f"Dept: {dept}    " + itxt2
                L(card, itxt2, C['text2'], 'sm').pack(anchor='w', padx=14, pady=(0,10))

        users_lb.bind('<<ListboxSelect>>', show_user)
        L(right, 'Select a user from the list to view their history.', C['text3'], 'body').pack(anchor='w', pady=40)


    def _e_staff_history(self):
        """Export all users' order and issue history to Excel"""
        if not self._xlok(): return
        f = self._ef()
        if not f: return
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
        wb = openpyxl.Workbook()
        gold = PatternFill('solid', fgColor='C9972C')
        navy = PatternFill('solid', fgColor='0B1F3A')
        thin = Side(style='thin', color='C9972C')
        bd2 = Border(left=thin, right=thin, top=thin, bottom=thin)
        ctr = Alignment(horizontal='center', vertical='center')
        bf2 = Font('Calibri', size=10, color='222222')
        dt_now = datetime.now().strftime('%d/%m/%Y %H:%M')

        # Sheet 1: Order History
        ws1 = wb.active; ws1.title = 'Order History'
        headers = ['User','Full Name','Role','Order#','Department','Status','Date','Item','Qty','Unit']
        n = len(headers)
        ws1.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n)
        c2=ws1.cell(1,1,'BSB STAFF ORDER HISTORY'); c2.fill=gold; c2.font=Font('Calibri',bold=True,color='0B1F3A',size=13); c2.alignment=ctr
        ws1.cell(2,1,f'Generated: {dt_now}').font=Font(italic=True,color='888888',size=10)
        for ci,h in enumerate(headers,1):
            c2=ws1.cell(3,ci,h); c2.fill=navy; c2.font=Font('Calibri',bold=True,color='FFFFFF',size=11); c2.alignment=ctr; c2.border=bd2
        rows=qy("""SELECT u.Username,ISNULL(u.FullName,''),u.Role,
                         o.OrderID,ISNULL(o.Department,''),o.Status,o.CreatedAt,
                         oi.ItemName,oi.Quantity,ISNULL(oi.Unit,'')
                   FROM Users u
                   JOIN Orders o ON o.OrderedBy=u.Username
                   JOIN OrderItems oi ON oi.OrderID=o.OrderID
                   ORDER BY u.FullName, o.CreatedAt DESC""")
        status_fills={'Pending':'FFF4E5','Arranged':'E7F7EE','Collected':'F3F6FC','OutOfStock':'FDECEC'}
        for ri,r in enumerate(rows,4):
            dt=r[6].strftime('%d/%m/%Y %H:%M') if hasattr(r[6],'strftime') else str(r[6])
            fill=PatternFill('solid',fgColor=status_fills.get(r[5],'FFFFFF'))
            vals=(r[0],r[1],r[2],r[3],r[4],r[5],dt,r[7],r[8],r[9])
            for ci,v in enumerate(vals,1):
                c2=ws1.cell(ri,ci,v); c2.fill=fill; c2.font=bf2; c2.border=bd2
        ws1.column_dimensions['B'].width=22; ws1.column_dimensions['H'].width=34

        # Sheet 2: Issue History
        ws2=wb.create_sheet('Issue History')
        h2=['Issued To','Department','Issued By','Date','Status','Item','Qty']
        n2=len(h2)
        ws2.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n2)
        c2=ws2.cell(1,1,'BSB STAFF ISSUE HISTORY'); c2.fill=gold; c2.font=Font('Calibri',bold=True,color='0B1F3A',size=13); c2.alignment=ctr
        for ci,h in enumerate(h2,1):
            c2=ws2.cell(3,ci,h); c2.fill=navy; c2.font=Font('Calibri',bold=True,color='FFFFFF',size=11); c2.alignment=ctr; c2.border=bd2
        irows=qy("""SELECT r.RecipientName,ISNULL(r.Department,''),r.IssuedBy,r.IssuedAt,
                         CASE WHEN r.Completed=1 THEN 'Completed' ELSE 'Remaining' END,
                         ii.ItemName,ii.Quantity
                   FROM IssueRecords r JOIN IssueItems ii ON r.IssueID=ii.IssueID
                   ORDER BY r.IssuedAt DESC""")
        for ri,r in enumerate(irows,4):
            dt=r[3].strftime('%d/%m/%Y %H:%M') if hasattr(r[3],'strftime') else str(r[3])
            fill=PatternFill('solid',fgColor='E7F7EE' if r[4]=='Completed' else 'FFF4E5')
            vals=(r[0],r[1],r[2],dt,r[4],r[5],r[6])
            for ci,v in enumerate(vals,1):
                c2=ws2.cell(ri,ci,v); c2.fill=fill; c2.font=bf2; c2.border=bd2
        ws2.column_dimensions['A'].width=22; ws2.column_dimensions['F'].width=34

        path=os.path.join(f,f'BSB_StaffHistory_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')
        wb.save(path)
        if messagebox.askyesno('Saved',f'Staff history exported:\n{path}\n\nOpen now?',parent=self):
            os.startfile(path)


    # ══ REPORT ═══════════════════════════════════════════════
    def _report(self):
        C=T()
        self._clear(); self.configure(fg_color=C['bg'])
        root=ctk.CTkFrame(self,fg_color=C['bg']); root.pack(fill='both',expand=True)
        self._topbar(root,'Analytics','Reports',lambda: self._reload(self._report))
        sc=ctk.CTkScrollableFrame(root,fg_color=C['bg']); sc.pack(fill='both',expand=True,padx=24,pady=18)
        L(sc,'Export Reports',C['text'],'h2').pack(anchor='w',pady=(0,14))
        cards=ctk.CTkScrollableFrame(sc,fg_color=C['bg'],orientation='horizontal',height=230); cards.pack(fill='x',pady=(0,20))
        for icon,title,desc,fn in [('📦','Inventory Report','Full stock list with values',self._e_inv),
                                    ('⚠️','Low Stock Alert','Items below minimum',self._e_low),
                                    ('📋','Issue Log','Complete issue history',self._e_iss),
                                    ('🏢','Dept Summary','Totals per department',self._e_dept),
                                    ('🛒','Orders Report','All staff orders & status',self._e_orders),
                                    ('📂','Export All Reports','All reports in one Excel file',self._e_all)]:
            card=ctk.CTkFrame(cards,fg_color=C['card'],corner_radius=12,border_width=1,border_color=C['border'],width=220,height=200)
            card.pack(side='left',padx=8); card.pack_propagate(False)
            L(card,icon,fk='disp').pack(pady=(22,8))
            L(card,title,C['text'],'body_b').pack()
            L(card,desc,C['text3'],'sm').pack(pady=(4,12),padx=14)
            blue_btn(card,'Export Excel',fn,w=180,h=38).pack()
        if self.current_role not in PROCUREMENT_ROLES:
            return
        hline(sc,12)
        L(sc,'Issue History',C['text'],'h3').pack(anchor='w',pady=(0,10))
        holder,ht=mktree(sc,('Date','Issued By','Recipient','Department','Items','Total BHD'),[150,130,160,140,70,120],height=12)
        holder.pack(fill='both')
        rows=qy("SELECT r.IssuedAt,r.IssuedBy,r.RecipientName,r.Department,COUNT(ii.IssueItemID),SUM(ii.Quantity*ii.UnitPrice) FROM IssueRecords r LEFT JOIN IssueItems ii ON r.IssueID=ii.IssueID GROUP BY r.IssueID,r.IssuedAt,r.IssuedBy,r.RecipientName,r.Department ORDER BY r.IssuedAt DESC")
        for i,r in enumerate(rows):
            dt=r[0].strftime('%d/%m/%Y %H:%M') if r[0] else ''
            total=f'{float(r[5]):.3f}' if r[5] else '0.000'
            ht.insert('','end',values=(dt,r[1],r[2],r[3],r[4],total),tags=('alt',) if i%2==0 else ())

    def _ef(self): return filedialog.askdirectory(title='Select folder',parent=self)
    def _xlok(self):
        try: import openpyxl; return True
        except: messagebox.showerror('Missing','Run: pip install openpyxl',parent=self); return False
    def _wb(self,sheet,headers,title):
        import openpyxl
        from openpyxl.styles import PatternFill,Font,Alignment,Border,Side
        wb=openpyxl.Workbook(); ws=wb.active; ws.title=sheet
        navy=PatternFill('solid',fgColor='0B1F3A'); gold=PatternFill('solid',fgColor='C9972C')
        thin=Side(style='thin',color='C9972C'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
        ctr=Alignment(horizontal='center',vertical='center'); n=len(headers)
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n)
        c=ws.cell(1,1,title); c.fill=gold; c.font=Font('Calibri',bold=True,color='0B1F3A',size=14); c.alignment=ctr
        ws.cell(2,1,f'Generated: {datetime.now().strftime("%d/%m/%Y %H:%M")}').font=Font(italic=True,color='888888',size=10)
        for ci,h in enumerate(headers,1):
            c=ws.cell(3,ci,h); c.fill=navy; c.font=Font('Calibri',bold=True,color='FFFFFF',size=11); c.alignment=ctr; c.border=bd
        return wb,ws,bd,Font('Calibri',size=10,color='222222')
    def _sv(self,wb,folder,fname):
        path=os.path.join(folder,fname); wb.save(path)
        if messagebox.askyesno('Saved',f'Saved!\n{path}\n\nOpen now?',parent=self): os.startfile(path)
    def _e_inv(self):
        if not self._xlok(): return
        f=self._ef()
        if not f: return
        from openpyxl.styles import PatternFill
        wb,ws,bd,bf=self._wb('Inventory',['#','Item Name','Category','Qty','Unit','Price','Total'],'BSB INVENTORY REPORT')
        rows=qy("SELECT ROW_NUMBER() OVER(ORDER BY s.Name,i.Name),i.Name,ISNULL(s.Name,'-'),i.Quantity,i.Unit,i.UnitPrice,i.Quantity*i.UnitPrice FROM Items i LEFT JOIN SubCategories s ON i.SubCatID=s.SubCatID ORDER BY s.Name,i.Name")
        for ri,r in enumerate(rows,4):
            fill=PatternFill('solid',fgColor='EAF1FE' if ri%2==0 else 'FFFFFF')
            for ci,v in enumerate(r,1):
                c=ws.cell(ri,ci,v); c.fill=fill; c.font=bf; c.border=bd
                if ci in(6,7): c.number_format='0.000'
        ws.column_dimensions['B'].width=34
        self._sv(wb,f,f'BSB_Inventory_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')
    def _e_low(self):
        if not self._xlok(): return
        f=self._ef()
        if not f: return
        from openpyxl.styles import PatternFill
        wb,ws,bd,bf=self._wb('Low Stock',['Item','Category','Qty','Min','Shortage','Price'],'BSB LOW STOCK ALERT')
        rows=qy("SELECT i.Name,ISNULL(s.Name,'-'),i.Quantity,i.MinStock,i.MinStock-i.Quantity,i.UnitPrice FROM Items i LEFT JOIN SubCategories s ON i.SubCatID=s.SubCatID WHERE i.Quantity<=i.MinStock ORDER BY(i.MinStock-i.Quantity) DESC")
        rf=PatternFill('solid',fgColor='FFF4E5')
        for ri,r in enumerate(rows,4):
            for ci,v in enumerate(r,1):
                c=ws.cell(ri,ci,v); c.fill=rf; c.font=bf; c.border=bd
        self._sv(wb,f,f'BSB_LowStock_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')
    def _e_iss(self):
        if not self._xlok(): return
        f=self._ef()
        if not f: return
        from openpyxl.styles import PatternFill
        wb,ws,bd,bf=self._wb('Issues',['Date','By','Recipient','Dept','Item','Qty','Price','Total'],'BSB ISSUE LOG')
        rows=qy("SELECT r.IssuedAt,r.IssuedBy,r.RecipientName,r.Department,ii.ItemName,ii.Quantity,ii.UnitPrice,ii.Quantity*ii.UnitPrice FROM IssueRecords r JOIN IssueItems ii ON r.IssueID=ii.IssueID ORDER BY r.IssuedAt DESC")
        for ri,r in enumerate(rows,4):
            fill=PatternFill('solid',fgColor='EAF1FE' if ri%2==0 else 'FFFFFF')
            for ci,v in enumerate(r,1):
                val=v.strftime('%d/%m/%Y %H:%M') if hasattr(v,'strftime') else v
                c=ws.cell(ri,ci,val); c.fill=fill; c.font=bf; c.border=bd
                if ci in(7,8): c.number_format='0.000'
        self._sv(wb,f,f'BSB_Issues_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')
    def _e_dept(self):
        if not self._xlok(): return
        f=self._ef()
        if not f: return
        from openpyxl.styles import PatternFill
        wb,ws,bd,bf=self._wb('Dept',['Department','Issues','Items','Total BHD'],'BSB DEPARTMENT SUMMARY')
        rows=qy("SELECT r.Department,COUNT(DISTINCT r.IssueID),SUM(ii.Quantity),SUM(ii.Quantity*ii.UnitPrice) FROM IssueRecords r JOIN IssueItems ii ON r.IssueID=ii.IssueID GROUP BY r.Department ORDER BY SUM(ii.Quantity*ii.UnitPrice) DESC")
        for ri,r in enumerate(rows,4):
            fill=PatternFill('solid',fgColor='EAF1FE' if ri%2==0 else 'FFFFFF')
            for ci,v in enumerate(r,1):
                c=ws.cell(ri,ci,v); c.fill=fill; c.font=bf; c.border=bd
                if ci==4: c.number_format='0.000'
        self._sv(wb,f,f'BSB_Dept_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')


    def _e_orders(self):
        if not self._xlok(): return
        f=self._ef()
        if not f: return
        import os
        from openpyxl.styles import PatternFill
        wb,ws,bd,bf=self._wb('Orders',['Order#','Ordered By','Department','Status','Date','Item','Qty','Unit'],'BSB ORDERS REPORT')
        rows=qy("""SELECT o.OrderID,o.OrderedByName,ISNULL(o.Department,''),o.Status,o.CreatedAt,
                         oi.ItemName,oi.Quantity,ISNULL(oi.Unit,'')
                   FROM Orders o JOIN OrderItems oi ON o.OrderID=oi.OrderID
                   ORDER BY o.CreatedAt DESC""")
        sc={'Pending':'FFF4E5','Arranged':'E7F7EE','Collected':'F3F6FC','OutOfStock':'FDECEC'}
        for ri,r in enumerate(rows,4):
            dt=r[4].strftime('%d/%m/%Y %H:%M') if hasattr(r[4],'strftime') else str(r[4])
            fill=PatternFill('solid',fgColor=sc.get(r[3],'FFFFFF'))
            vals=(r[0],r[1],r[2],r[3],dt,r[5],r[6],r[7])
            for ci,v in enumerate(vals,1):
                c2=ws.cell(ri,ci,v); c2.fill=fill; c2.font=bf; c2.border=bd
        ws.column_dimensions['F'].width=32
        self._sv(wb,f,f'BSB_Orders_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')

    def _e_all(self):
        if not self._xlok(): return
        f=self._ef()
        if not f: return
        import os, openpyxl
        from openpyxl.styles import PatternFill,Font,Border,Side,Alignment
        wb=openpyxl.Workbook()
        gold=PatternFill('solid',fgColor='C9972C'); navy=PatternFill('solid',fgColor='0B1F3A')
        thin=Side(style='thin',color='C9972C'); bd2=Border(left=thin,right=thin,top=thin,bottom=thin)
        ctr=Alignment(horizontal='center',vertical='center')
        bf2=Font('Calibri',size=10,color='222222')
        def add_sheet(title,headers,rows_data,sheet_name):
            ws=wb.create_sheet(sheet_name)
            n=len(headers)
            ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n)
            c2=ws.cell(1,1,title); c2.fill=gold; c2.font=Font('Calibri',bold=True,color='0B1F3A',size=13); c2.alignment=ctr
            for ci,h in enumerate(headers,1):
                c2=ws.cell(3,ci,h); c2.fill=navy; c2.font=Font('Calibri',bold=True,color='FFFFFF',size=11); c2.alignment=ctr; c2.border=bd2
            for ri,r in enumerate(rows_data,4):
                fill=PatternFill('solid',fgColor='EAF1FE' if ri%2==0 else 'FFFFFF')
                for ci,v in enumerate(r,1):
                    c2=ws.cell(ri,ci,v); c2.fill=fill; c2.font=bf2; c2.border=bd2
        inv=qy("SELECT ROW_NUMBER() OVER(ORDER BY s.Name,i.Name),i.Name,ISNULL(s.Name,'-'),i.Quantity,i.Unit,i.UnitPrice,i.Quantity*i.UnitPrice FROM Items i LEFT JOIN SubCategories s ON i.SubCatID=s.SubCatID ORDER BY s.Name,i.Name")
        add_sheet('BSB INVENTORY',['#','Item','Category','Qty','Unit','Price','Total'],inv,'Inventory')
        low=qy("SELECT i.Name,ISNULL(s.Name,'-'),i.Quantity,i.MinStock,i.MinStock-i.Quantity,i.UnitPrice FROM Items i LEFT JOIN SubCategories s ON i.SubCatID=s.SubCatID WHERE i.Quantity<=i.MinStock ORDER BY(i.MinStock-i.Quantity) DESC")
        add_sheet('BSB LOW STOCK',['Item','Category','Qty','Min','Shortage','Price'],low,'Low Stock')
        iss=qy("SELECT r.IssuedAt,r.IssuedBy,r.RecipientName,r.Department,ii.ItemName,ii.Quantity,ii.UnitPrice FROM IssueRecords r JOIN IssueItems ii ON r.IssueID=ii.IssueID ORDER BY r.IssuedAt DESC")
        iss2=[(r[0].strftime('%d/%m/%Y') if hasattr(r[0],'strftime') else r[0],)+r[1:] for r in iss]
        add_sheet('BSB ISSUE LOG',['Date','By','Recipient','Dept','Item','Qty','Price'],iss2,'Issue Log')
        dept=qy("SELECT r.Department,COUNT(DISTINCT r.IssueID),SUM(ii.Quantity),SUM(ii.Quantity*ii.UnitPrice) FROM IssueRecords r JOIN IssueItems ii ON r.IssueID=ii.IssueID GROUP BY r.Department ORDER BY SUM(ii.Quantity*ii.UnitPrice) DESC")
        add_sheet('BSB DEPT SUMMARY',['Department','Issues','Items','Total BHD'],dept,'Dept Summary')
        ords=qy("SELECT o.OrderID,o.OrderedByName,ISNULL(o.Department,''),o.Status,o.CreatedAt,oi.ItemName,oi.Quantity,ISNULL(oi.Unit,'') FROM Orders o JOIN OrderItems oi ON o.OrderID=oi.OrderID ORDER BY o.CreatedAt DESC")
        ords2=[(r[0],r[1],r[2],r[3],r[4].strftime('%d/%m/%Y') if hasattr(r[4],'strftime') else r[4],r[5],r[6],r[7]) for r in ords]
        add_sheet('BSB ORDERS',['Order#','By','Department','Status','Date','Item','Qty','Unit'],ords2,'Orders')
        if 'Sheet' in wb.sheetnames: del wb['Sheet']
        path=os.path.join(f,f'BSB_AllReports_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')
        wb.save(path)
        if messagebox.askyesno('Saved',f'All reports saved:\n{path}\n\nOpen now?',parent=self):
            os.startfile(path)


if __name__ == '__main__':
    app = BSBApp()
    app.mainloop()