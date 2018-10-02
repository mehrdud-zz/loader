import os
import json
import logging
import datetime
import sys
import file_structure as file_structure
import excel_openpyxl as excel_openpyxl 
from os import listdir
from os.path import isfile, join

excel_loader = "openpyxl"
 
def create_config_file_structure_information_file(rebalance_tool_inputs, logger):
    files = []
    for file in rebalance_tool_inputs["RebalancerToolInputs"]:
        if(file["Type"]=="file"):
            headers = []
            headers = excel_reader().get_headers(file["Path"], 0,logger)    
            file= {"Headers": headers, "Index": len(files), "Path": file["Path"], "File": file["File"]}
            files.append(file)
    return files
    
def setup_logger(name, log_file, level=logging.DEBUG):
    formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
    handler = logging.FileHandler(log_file)                  
    handler.setFormatter(formatter)
    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.addHandler(handler)
    return logger

def excel_reader():    
    return excel_openpyxl

def run_loader(domain_configuration_file_path):
    with open(domain_configuration_file_path, 'r') as domain_configuration_file:
        domain_configurations = json.load(domain_configuration_file)
        datetime_string = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        for domain in domain_configurations["Domains"]:
            folders = file_structure.create_folder_structure(domain["OutputFolder"], datetime_string,domain["Domain"])
            log_file_name = folders["Logs"] + domain["Domain"]+ "_Log.log" 		
            domain_logger = setup_logger(domain["Domain"] + "_Logger", log_file_name, logging.DEBUG)
            rebalance_tool_config_file_list = file_structure.compile_list_of_files(domain["ConfigFileList"])
            rebalance_tool_config_file_list_structured = create_config_file_structure_information_file(rebalance_tool_config_file_list, domain_logger)
            result = json.dumps(rebalance_tool_config_file_list_structured, indent=2, sort_keys=True)
            file_structure.write_file(folders["Configuration"] + "\\Structured_Configs.json", result)
            for config_file in rebalance_tool_config_file_list_structured:
                  print("File: " + config_file["File"] + ", Path: " + config_file["Path"])
                  log_file_name = folders["Logs"] + config_file["File"] + ".log"
                  logger = setup_logger(config_file["File"] + "_Logger", log_file_name)
                  config_file_content = excel_reader().read_worksheet_content(config_file["Path"], config_file["Headers"], 0 , logger)
                  print("    Retrieved " + str(len(config_file_content)) + " rows \r\n\r\n")
                  file_structure.write_file(folders["Output"] + config_file["File"]+".json", json.dumps(config_file_content, indent=2, sort_keys=True))

def start():
    domain_configuration_file_path = "\\\\lon0306.london.schroders.com\\dfs\\home3\\users\\nateghm\\My Documents\\Projects\\Multi-Asset Core Platform\\DataLoader\\Configs\\Configuration.json" 
    run_loader(domain_configuration_file_path)
    return "Started job"

start()