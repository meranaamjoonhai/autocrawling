import schedule
import time
import subprocess

def job():
    subprocess.run(["python", "crawling.py"])

schedule.every().day.at("15:13").do(job)

while True:
    schedule.run_pending()
    time.sleep(1)