# IMPORTING LIBRARIES
from bs4 import BeautifulSoup
import requests
import csv
import re
import pandas as pd
import warnings
from requests.adapters import HTTPAdapter
warnings.filterwarnings("ignore")
from googlesearch import search
from datetime import datetime
import os
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

# METHOD TO CLEAN THE CONTENT OF THE PAGE
def cleaning(content):
    clean = re.compile("<script(.|\n)*?>(.|\n)*?</script>")
    text = re.sub(clean, ' ', str(content))

    clean = re.compile("<style(.|\n)*?>(.|\n)*?</style>")
    text = re.sub(clean, ' ', text)

    tags = re.compile("<(.|\n)*?>")
    text = re.sub(tags, ' ', text) # THIS IS THE CLEAN TEXT
    return text

# METHOD TO GET THE RESULTS FROM GOOGLE
def get_content_from_google(institution):
    search_results = search(institution, num=5, stop=5, pause=2.0)
    
    ignore_list = ["facebook", "instagram", "moneyhouse", "monitor", "linkedin", 
                   "dnb", "opencorporates", "kompass", "tel", "register", "local", "business"]
    # STORING THE RESULTS IN AN ARRAY
    for result in search_results:
        if any(i in str(result) for i in ignore_list):
            continue
        else:
            return str(result).split('/')[0] + "//" + str(result).split('/')[2]
        
    return None

# GETTING THE CONTENT OF THE WEBSITE
def get_content(url, institution):
    except_flag = False
    try: 
        r = requests.get(url[0], verify=False)
        ret_url = url[0]
    except Exception as e:
        except_flag = True
        print(str(url[0]) + " not found!")
    finally:
        if except_flag == True or (r.status_code != 200 or len(cleaning(r.content)) < 30):
            try: 
                except_flag = False
                r = requests.get(url[1], verify=False)
                ret_url = url[1]
            except Exception as e:
                except_flag = True
                print(str(url[1]) + " not found!")
            finally:
                if except_flag or r.status_code != 200 or len(cleaning(r.content)) < 30:
                    ret_url = get_content_from_google(institution)
                    print("Given URL was not found. This URL was found from Google:", ret_url)
                    if ret_url == None:
                        print(str(ret_url) + " not found!")
                        return None, None  
                    try: 
                        r = requests.get(ret_url, verify=False)
                    except Exception as e:
                        print(str(ret_url) + " not found!")
                        return None, None
        return r, ret_url

# METHOD TO CHECK THE CONTENTS OF THE URL
def check_content(r, dictionary, info_to_find):
    soup = BeautifulSoup(r.content) # If this line causes an error, run 'pip install html5lib' or install html5lib 
    a_tags = soup.findAll('a', href=True)

    # CLEANING THE CONTENTS OF THE PAGE
    text = cleaning(r.content)
    
    # CHECKING THE CONTENTS OF THE PAGE
    for key,info in zip(list(dictionary.keys()),info_to_find):
        # IF THE CONTENT IS ALREADY FOUND, IT WONT BE CHECKED AGAIN
        if dictionary[key] == True:
            continue
        if str(info) in text:
            dictionary[key] = True
    
    # GETTING THE EMAILS ON THE PAGE
    email_list = get_emails(text) 

    return dictionary, email_list

# SAVING ALL THE LINKS IN AN ARRAY
def get_links(r, url):
    # DEFINING BEAUTIFUL SOUP
    soup = BeautifulSoup(r.content) # If this line causes an error, run 'pip install html5lib' or install html5lib 
    
    # GETTING ALL THE LINKS FROM THE HOME PAGE
    a_tags = soup.findAll('a', href=True)
    links = []
    for tag in a_tags:
        if 'http' in tag['href']:
            if url in tag['href']:
                links.append(tag['href'])
        else:
            if 'mail' not in tag['href']:
                if 'index' not in tag['href']:
                    links.append(url + "/" + tag['href'])
    return links

# GETTING THE EMAILS OF A PAGE
def get_emails(text):
    regex = re.compile('[\w\.\-]+@[\w\.\-]+\.[\w]{2,4}')
    emails = re.findall(regex, text)

    return emails

# CHECKING IF ANY EMAIL BELONGS TO THE NAME
def check_emails(email_list, name):
    emails_found = []
    for email in email_list:
        
        for n in name.split(' '):            
            if n.lower() in email.split('@')[0]:
                emails_found.append(email)
        
        if (name.split(' ')[0][0] + name.split(' ')[-1][0]).lower() in email.split('@')[0]:
            emails_found.append(email)
    
    return emails_found

# CHECKING IF ALL THE VALUES IN THE DICTIONARY ARE FOUND OR NOT
def check_dictionary(dictionary):
    for value in list(dictionary.values()):
        if value == False:
            return False
    return True

# RUNNER METHOD
def runner(file_name, output_file):
    # READING FROM A FILE
    df = pd.read_excel(file_name, engine='openpyxl')

    prev_url = ''
    prev_dictionary = {}
    prev_emails = []

    output_df = pd.DataFrame(columns=['Time Stamp', '', 'Institution', 'Status', 'Strasse', 'PLZ', 'Ort', 'Kanton', 'Rechtsform', 'Industrie', 'Noga', 'Employee Size', 'Website', 'Beschreibung', 'Ziel', 'UID', 'Telefonnummer', 'Allgemeine Email', 'Entscheidungsträger', 'Funktion', 'Persönliche E-Mail', 'URL', 'Email list', 'Linked Emails'] )

    # ITERATING THROUGH EVERY ROW OF THE XLSX FILE
    for count, row in df.iterrows():
        output_array = list(row)

        # information to find on page
        info_to_find = row['Strasse'],row['PLZ'],row['Ort'],row['Telefonnummer'], row['Allgemeine Email'],row['Entscheidungsträger'],row['Persönliche E-Mail']
        info_to_find = list(info_to_find)
        
        # NEED SOME PRE PROCESSING FOR TELEFONE NUMBER
        if str(info_to_find[3]) != 'nan':
            info_to_find[3] = info_to_find[3][1:]
        emails_found = []
        email_list = []
        dictionary = {
                        'Strasse': False, 
                        'PLZ': False, 
                        'Ort': False, 
                        'Telefonnummer': False, 
                        'Allgemeine Email': False, 
                        'Entscheidungsträger': False,  
                        'Persönliche E-Mail': False
        }

        if row['Website'].split('/')[0] == prev_url:
            dictionary = prev_dictionary
            email_list = prev_emails

            if str(info_to_find[5]) != 'nan':
                emails_found = check_emails(email_list, info_to_find[5])

            # REMOVING THE ELEMENTS WHICH WERE NOT FOUND
            for ind, header in enumerate(df.columns):
                if header in dictionary.keys():
                    if dictionary[header] == False:
                        output_array[ind] = ""

            # ADDING TIMESTAMP TO THE OUTPUT ARRAY
            time = datetime.now().strftime('%d/%m/%Y, %H:%M')
            output_array.insert(0, time)

            # ADDING EMAIL LIST TO THE OUTPUT ARRAY
            output_array.append(",".join(email_list))

            # ADDING LINKED EMAILS TO THE OUTPUT ARRAY
            output_array.append(",".join(emails_found))

            output_array[12] = parent_url
            output_df.loc[-1] = output_array
            output_df.index += 1

        else:
            prev_url = row['Website'].split('/')[0]
            url = ['https://' + prev_url, 'http://' + prev_url] # APPENDING PROTOCOL TO THE WEBISTE NAME
            
            print("URL:", url)
            
            # GETTING THE CONTENTS
            r, url = get_content(url, row['Institution'])

            # SAVING THE PARENT URL TO BE USED LATER
            parent_url = url
            # IF URL IS NOT FOUND, WE MOVE ON TO THE NEXT ROW
            if r == None:
                continue
            
            # CHECKING THE CONTENTS OF THE URL
            dictionary, emails = check_content(r, dictionary, info_to_find)
            
            # GETTING THE EMAILS IN THE PAGE
            for email in emails:
                email_flag = False
                if email not in email_list:
                    for l_email in email_list:
                        if l_email in email:
                            email_flag = True
                            break
                    if not email_flag:
                        email_list.append(email)
            # GETTING THE LINKS OF THE HOMEPAGE
            links = get_links(r, url)

            print(url)

            for link in links:
                link_array = [link, None]
                # OPTIMIZING THE CODE
                if check_dictionary(dictionary):
                    break
                
                # GETTING CONTENT
                r, url = get_content(link_array, row['Institution'])
                
                if r == None:
                    continue

                # CHECKING THE CONTENTS OF THE URL
                dictionary, emails = check_content(r, dictionary, info_to_find)
                
                # GETTING THE EMAILS IN THE PAGE 
                for email in emails:
                    email_flag = False
                    if email not in email_list:
                        for l_email in email_list:
                            if l_email in email:
                                email_flag = True
                                break
                        if not email_flag:
                            email_list.append(email)
                
                if str(info_to_find[5]) != 'nan':
                    emails_found = check_emails(email_list, info_to_find[5])

            prev_dictionary = dictionary
            prev_emails =  email_list

            # REMOVING THE ELEMENTS WHICH WERE NOT FOUND
            for ind, header in enumerate(df.columns):
                if header in dictionary.keys():
                    if dictionary[header] == False:
                        output_array[ind] = ""

            # ADDING TIMESTAMP TO THE OUTPUT ARRAY
            time = datetime.now().strftime('%d/%m/%Y, %H:%M')
            output_array.insert(0, time)

            # ADDING EMAIL LIST TO THE OUTPUT ARRAY
            output_array.append(",".join(email_list))

            # ADDING LINKED EMAILS TO THE OUTPUT ARRAY
            output_array.append(",".join(emails_found))

            output_array[12] = parent_url
            output_df.loc[-1] = output_array
            output_df.index += 1

            print("Information log:\n", dictionary) 
            print("Email list:\n", email_list)
        print("Entscheidungsträger:", row['Entscheidungsträger'])
        print("Emails found: ", emails_found)
        print()

        # SAVING THE OUPUT TO A CSV FILE AFTER 20 ITERATIONS
        if len(output_df) % 20 == 0:
            output_df.to_csv(output_file, encoding="utf-8", index=False)

    return output_df

# RUNNER 
path = 'active/'
entries = os.listdir("active/")

for entry in entries:
    print("Input file:", entry)

    output_file = 'output/output-file-' + entry.split('.')[0]+ '.csv'
    output = runner(path + entry, output_file)

    output.to_csv(output_file, encoding="utf-8", index=False)

    print("Output file created:", output_file)
    print()