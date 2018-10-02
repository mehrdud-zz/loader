import os
import glob
import logging
import json

def create_folder_structure(path, jobId, domain):
   response_folders = {
   "Main": path + "\\" + jobId,
   "Configuration" : path + "\\" + jobId + "\\" + domain + "\\" + "Configuration\\",
   "Logs": path + "\\" + jobId + "\\" + domain + "\\" + "Logs\\",
   "Output": path + "\\" + jobId + "\\" + domain + "\\" + "Output\\"
   }

   if not os.path.exists(response_folders["Main"]):
      os.mkdir(response_folders["Main"])
      os.mkdir(response_folders["Main"] + "\\" + domain)
      os.mkdir(response_folders["Configuration"])
      os.mkdir(response_folders["Logs"])
      os.mkdir(response_folders["Output"])
   return response_folders
    
    
def compile_list_of_files(input_list):		   
   with open(input_list, 'r') as input_file:
      input_list_json_data = json.load(input_file)
      for item in input_list_json_data["RebalancerToolInputs"]:
         if(item["Type"] == "folder"):
            list_of_files = glob.glob(item['Path'] + '\\*.xlsx', recursive=True) 
            if(not not list_of_files):
               latest_file = max(list_of_files, key=os.path.getctime)
               file_json_data = {"File": os.path.basename(latest_file), "Path": latest_file, "Type": "file"}
               input_list_json_data["RebalancerToolInputs"].append(file_json_data)
      return input_list_json_data


def write_file(file_path, content):
   with open(file_path, "w") as text_file:
      print(content, file=text_file)