import schedule
import time
import subprocess

def job():
    subprocess.run(["python", "crawling.py"])

schedule.every().day.at("16:55").do(job)


while True:
    schedule.run_pending()
    time.sleep(1)
