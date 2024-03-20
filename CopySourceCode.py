from argparse import ArgumentParser
import logging
import os
import shutil
import openpyxl
from datetime import datetime

def read_excel_data(file_path):
    data = []
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            data.append((row[0], row[4]))
        workbook.close()
    except Exception as e:
        print(f"Error reading Excel file: {e}\n")
        logging.error(f"Error reading Excel file: {e}\n")
    return data

def create_directory_if_not_exists(directory_path):
    try:
        if not os.path.exists(directory_path):
            try:
                os.makedirs(directory_path)
                print(f"Directory '{directory_path}' created successfully.\n")
                logging.info(f"Directory '{directory_path}' created successfully.\n")
            except OSError as e:
                print(f"Error: {e}\n")
                logging.error(f"Error: {e}\n")
        else:
            print(f"Directory '{directory_path}' already exists.\n")
            logging.info(f"Directory '{directory_path}' already exists.\n")
    except Exception as e:
        print(f"Error while creating a directory: {e}\n")
        logging.error(f"Error while creating a directory: {e}\n")

def move_and_delete_folders(root_dir):
    try:
        # Iterate over subdirectories recursively
        for subdir, dirs, files in os.walk(root_dir):
            for folder in dirs:
                if folder.startswith('lmigtech-'):
                    source_dir = os.path.join(subdir, folder)
                    destination_dir = os.path.dirname(source_dir)
                    # Move contents of source directory to parent directory
                    for item in os.listdir(source_dir):
                        item_path = os.path.join(source_dir, item)
                        if os.path.isfile(item_path):
                            shutil.move(item_path, destination_dir)
                        elif os.path.isdir(item_path):
                            shutil.move(item_path, os.path.join(destination_dir, item))
                    # Remove the now empty source directory
                    os.rmdir(source_dir)
    except Exception as e:
        print(f"Error while moving and deleting folders: {e}\n")
        logging.error(f"Error while moving and deleting folders: {e}\n")

def copy_source_code(data, input_path, output_path):
    try:
        repos_and_apps = []
        apps = []
        for item in data:
            repos_and_apps.append((item[0],item[1]))
            if item[1] not in apps:
                apps.append(item[1])
                application_name_directory = output_path+'\\'+item[1]
                create_directory_if_not_exists(application_name_directory)

        for item in repos_and_apps:
            repo = item[0]
            app = item[1]
            input_repo = os.path.join(input_path, repo)
            output_repo = os.path.join(output_path, app, repo)
            if os.path.exists(output_repo):
                logging.info(f'Repo - "{repo}" is already copied to the Path - "{output_path}"\n')
                print(f'Repo - "{repo}" is already copied to the Path - "{output_path}"\n')
            elif os.path.exists(input_repo):
                shutil.copytree(input_repo, output_repo)
                logging.info(f'Repo - "{repo}" is copied successfully from "{input_path}" to "{output_path}".\n')
                print(f'Repo - "{repo}" is copied successfully from "{input_path}" to "{output_path}".\n')
            elif not os.path.exists(input_repo):
                logging.info(f'Repo - "{repo}" is not available inside the Path - "{input_path}"\n')
                print(f'Repo - "{repo}" is not available inside the Path - "{input_path}"\n')

    except Exception as e:
        print(f"Error while copying source code: {e}\n")
        logging.error(f"Error while copying source code: {e}\n")


if __name__ == "__main__":

    parser = ArgumentParser()
 
    parser.add_argument('-excel_file', '--excel_file', required=True, help='Excel File Name')
    parser.add_argument('-input_path','--input_path', required=True, help='Input Path')
    parser.add_argument('-output_path', '--output_path', required=True, help='Output Path')

    args=parser.parse_args()

    os.makedirs('CopySourceCode_logs', exist_ok=True)
    datetime_now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    # Set up logging
    log_file = os.path.join('CopySourceCode_logs',f"CopySourceCode_logs_{datetime_now}.log")
    logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


    data = read_excel_data(args.excel_file)
    # print(data)

    copy_source_code(data, args.input_path, args.output_path)
    move_and_delete_folders(args.output_path)