import datetime 
import Validations as v
import os
from shutil import copy

def main():
    try : 
        #converting date format from 2024-04-15 to 20240415
        today_date = datetime.date.today().strftime('%Y%m%d')
        
        #paths
        master_file_path=f"../Master_Data"
        success_file_path=f'../../Success_files/{today_date}'
        rejected_file_path=f'../../Rejected_files/{today_date}'
        incoming_file_path=f'../Incoming_files/{today_date}'
        
        #getting list of files to incoming_files
        incoming_files=os.listdir(incoming_file_path)
        #email_date = datetime.date.today().strftime('%Y-%m-%d')
        #subject = f'Validation email for {email_date}'
            
        #Validation starts if files existed
        if len(incoming_files)>0:
            #print(741)
            #os.chdir(incoming_file_path)
            
            #Cleaning temp files
            for x in incoming_files:
                if x[0]=='~':
                    print(582)
                    incoming_files.remove(x)
                    
            # To count the failed files
            failed=0
            
            #To count no. of files validated
            count=0
            print("Validation Started .....")
            #Now going to validate each file
            for file in incoming_files: 
                
                if v.validations(incoming_file_path,file,rejected_file_path):#validation done here
                    #print(741)
                    #if validation fails there we will be here
                    if not os.path.exists(f'{rejected_file_path}'):# checking path existence
                        os.makedirs(f'{rejected_file_path}', exist_ok=True)
                    
                    #we are moving to incoming files path to call file in shutil copy
                    os.chdir(f'../'+incoming_file_path)
                    #print(478)
                    copy(file, f'{rejected_file_path}/{file}')
                    #print(900)
                    failed+=1
                    
                else:
                    #If validation succeed then it enter here
                    if not os.path.exists(f'{success_file_path}'):
                        os.makedirs(f'{success_file_path}', exist_ok=True)
                    copy(file, f'{success_file_path}/{file}')
                count+=1    
                
                #prints no. of files need to validate           
                if count in range(10,len(incoming_files),10):
                    print(len(incoming_files)-count," need to validate\n")
                    
            #Finally it returns total files and success & failed ones
            #else will execute once all files would be validated                        
            else:
                print("Total Files",len(incoming_files),"\nSuccess : ",len(incoming_files)-failed,"\nFailed : ",failed)
        
        #this else will be execute if no files existed in incoming file path
        else:
            print("No file present in source folder.")
    
    #returns if any error raised       
    except Exception as msg:
        print('error :',msg)

#calling main function      
main()


