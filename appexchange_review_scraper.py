import re
from bs4 import BeautifulSoup
import requests
import json
import os
import xlsxwriter


def getJson(url):
    profilePage = requests.get(url).text
    soup = BeautifulSoup(profilePage, 'lxml')
    scriptTagWithProfileDataJson = soup.find_all("script")[17].string
    match = re.search('var profileData = JSON.parse(.*?);',
                      scriptTagWithProfileDataJson)
    extractedProfileData = match.group(0)
    # print(extractedProfileData)

    indexOfJSON = extractedProfileData.index("\"")
    # print(indexOfJSON)
    jsonData = extractedProfileData[indexOfJSON:]
    finalJSONStringWithSomeInvalidString = jsonData[:-2]
    # print(finalJSONString)

    jsonString = finalJSONStringWithSomeInvalidString.replace(
        "d\\'Ivoire", "dIvoire").replace("People\\'s", "Peoples")
    finalJson = json.loads(json.loads(jsonString))
    # print(test)
    # print(decodedJSON)
    return finalJson


def writeToExcel(url, finalJson, worksheet, index):

    user = finalJson["profileUser"]
    firstName = str(user['FirstName'])
    lastName = str(user['LastName'])
    profile = 'public'
    if 'Title' in user and user['Title']:
        title = str(user['Title'])
    else:
        title = str(None)
        profile = 'restricted'
    if 'CompanyName' in user and user['CompanyName']:
        company = str(user['CompanyName'])
    else:
        company = str(None)

    print("User - "+str(user))
    print("FName - "+firstName)
    print("LName - "+lastName)
    print("Title -"+title)
    print("Company -"+company)

    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 20)

    # Write some simple text.
    worksheet.write('A'+str(index), firstName)

    # Text with formatting.
    worksheet.write('B'+str(index), lastName)

    # Text with formatting.
    worksheet.write('C'+str(index), firstName + ' '+lastName)

    worksheet.write('D'+str(index), title)

    worksheet.write('E'+str(index), company)

    worksheet.write('F'+str(index), url)

    worksheet.write('G'+str(index), profile)


directory = "data"
baseURL = "https://trailblazer.me/id/"
url = ""

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('leadsfromappexchange1.xlsx')


for filename in os.listdir(directory):
    f = os.path.join(directory, filename)

    # checking if it is a file

    if os.path.isfile(f):
        print("------------------------")
        print(f)
        print("************************")
        with open(f) as fp:
            reviews = json.load(fp)
            responses = reviews['actions'][0]['returnValue']['returnValue']['responses']
            print("Number of reviews in the listing  -> " + str(len(responses)))

            worksheet = workbook.add_worksheet(str((os.path.basename(f))))
            for i in range(1, 3):
                respondedUser = responses[i]['responderUser']
                if respondedUser:
                    if 'trailblazerIdentityId' not in respondedUser:
                        continue
                    id = respondedUser['trailblazerIdentityId']
                    url = str(baseURL) + str(id)
                    print(str(url))
                    try:
                        writeToExcel(url, getJson(url), worksheet, i)
                    except:
                        continue


workbook.close()
