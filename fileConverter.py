import requests
import json
import sys
import os
import time

api_key = os.getenv('FILE_CONVERTER_API')
base_url = "https://api.freeconvert.com/v1"

url = ""

def wait_for_job_by_polling(job_id):
    for _ in range(10):
        time.sleep(2)
        job_get_response = freeconvert.get(f"{base_url}/process/jobs/{job_id}")
        job = job_get_response.json()

        if job["status"] == "completed":

            export_task = next(task for task in job["tasks"] if task["name"] == "myExport1")
            url = export_task["result"]["url"]
            print(url)
            return


        elif job["status"] == "failed":
            raise Exception("Job failed")

    raise Exception("Poll timeout")

if api_key == None:
    sys.exit("api key not found")

if len(sys.argv) != 2:
    upload_file_path = sys.argv[1]
    export_file_type = sys.argv[2]
else:
    sys.exit("Invalid number of arguments")

file_extension = os.path.splitext(upload_file_path)


freeconvert = requests.Session()
freeconvert.headers.update({
    "Content-Type": "application/json",
    "Accept": "application/json",
    "Authorization": f"Bearer {api_key}",
})


# Create an import/upload task
upload_task_response = freeconvert.post(f"{base_url}/process/import/upload")
upload_json = upload_task_response.json()
if "id" not in upload_json:
    print("Upload task response did not contain 'id'. Raw response:")
    print(json.dumps(upload_json, indent=2))
    sys.exit("Exiting due to an error.")

upload_task_id = upload_json["id"]
uploader_form = upload_json["result"]["form"]
# print("Created task", upload_task_id)

# Attach required form parameters and the file
formdata = {}
for parameter, value in uploader_form["parameters"].items():
    formdata[parameter] = value
files = {'file': open(upload_file_path,'rb')}

# Submit the upload as multipart/form-data request.
upload_response = requests.post(uploader_form["url"], files=files, data=formdata)
upload_response.raise_for_status()

# Use the uploaded file in a job.
# Job will complete when all its children and dependent tasks are complete.
job_response = freeconvert.post(f"{base_url}/process/jobs", json={
    "tasks": {
        "myConvert1": {
            "operation": "convert",
            "input": upload_task_id,
            "output_format": export_file_type,
        },
        "myExport1": {
            "operation": "export/url",
            "input": "myConvert1",
            "filename": "converted-file",
        },
    },
})
job = job_response.json()

if "id" not in job:
    print("Job response did not contain 'id'. Raw response:")
    try:
        print(json.dumps(job, indent=2))
    except Exception:
        print(repr(job))
    raise SystemExit(1)

# print("Job created", job["id"])

# Job will proceed as soon as the upload is finished.
# Need to wait for job completion/failure using polling or websocket (see relevant code examples).
wait_for_job_by_polling(job["id"])


