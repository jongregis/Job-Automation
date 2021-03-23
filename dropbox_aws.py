import dropbox
import openpyxl as xl
import boto3
import botocore

dbx = dropbox.Dropbox(
    'ucQp1NoOMzUAAAAAAAAAAXYCaTDU29D37vRXCkCwyQ0ep9kcdLbvHFjExMYzesBT')
s3 = boto3.resource('s3')
# for x in dbx.files_list_folder('/ECA Back Office/JON').entries:
#     print(x.name)
data = "Potential headline: Game 5 a nail-biter as Warriors inch out Cavs"
overwrite = True
mode = (dropbox.files.WriteMode.overwrite
        if overwrite
        else dropbox.files.WriteMode.add)
# dbx.files_upload(data.encode('UTF-8'),
#                  '/ECA Back Office/JON/story.xlsx', mode)

# pete = "/ECA Back Office/Pete's Backup/MILTARY/PETE ALL 3 SPREADSHEETS MYCAA FOR STACEY AND LISA/MAIN ENROLLMENT FOLDER/SPREADSHEETS/students mycaa FINAL-TODAY.xlsx"
# try:
#     s3.Bucket('jobautomation').download_file(
#         'Sept 2020.xlsx', '/Users/jongregis/Desktop/Sept 2020.xlsx')
# except botocore.exceptions.ClientError as e:
#     if e.response['Error']['Code'] == "404":
#         print("The object does not exist.")
#     else:
#         raise


def download_from_DB_to_S3(dbx, path, name):

    try:
        metadata, res = dbx.files_download(path=path)
        # f.write(res.content)
        s3.Bucket('jobautomation').put_object(
            Key=name, Body=res.content)

    except dropbox.exceptions.HttpError as err:
        print('*** HTTP error', err)
        return None

    print(f'Finished transferring {name} from DB to S3')


def main(month):
    monthly = f"/ECA Back Office/Lisa's Backup/Invoices/2020 Enrollment/{month} 2020.xlsx"
    download_from_DB_to_S3(dbx, pete, "students mycaa FINAL-TODAY.xlsx")
    download_from_DB_to_S3(dbx, monthly, f"{month} 2020.xlsx")


# download(dbx, path2, "jon")
# pete = "/Users/jongregis/Python/JobAutomation/pete.xlsx"
# wb1 = xl.load_workbook(pete)
# auburn = wb1.worksheets[0]
# print(auburn.max_row)
