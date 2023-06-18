# This is a sample Python script.
import smtplib
import subprocess
from dataclasses import dataclass
from email.mime.text import MIMEText
from typing import List

import wmi
from dataclass_wizard import YAMLWizard


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


@dataclass
class SenderConfig(YAMLWizard):
    email: str
    password: str


@dataclass
class Config(YAMLWizard):
    process: str
    alert_emails: List[str]
    exe_file_location: str
    auto_restart: str
    sender: SenderConfig


def run_exe_file(fp: str):
    print('Running file %s', fp)
    process = subprocess.Popen(fp, stdout=subprocess.PIPE, creationflags=0x08000000)


def send_email(sender: str, password: str, to: List[str], auto_restart: str):
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
        msg = MIMEText("Alert for CNGG syncing service has been down")
        msg['Subject'] = f"CNGG syncing down, auto restart: {auto_restart.upper()}"
        msg['From'] = sender
        msg['To'] = ', '.join(to)
        smtp_server.login(sender, password)
        smtp_server.sendmail(sender, to, msg.as_string())


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    with open('config.yaml', 'r') as fd:
        cfg = Config.from_yaml(fd.read())
    w = wmi.WMI()
    running = False
    processes = []
    for process in w.Win32_Process():
        processes.append(process.Name)
        if cfg.process in process.Name:
            running = True
    if not running:
        if cfg.auto_restart == 'enabled':
            run_exe_file(cfg.exe_file_location)
        send_email(
            cfg.sender.email,
            cfg.sender.password,
            cfg.alert_emails,
            cfg.auto_restart)
