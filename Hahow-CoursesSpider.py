import requests as req
from openpyxl import Workbook

# # Create an Excel
wb = Workbook()
# # Default sheet
ws = wb.active

# # Fill in the title on the sheet
title = ['Course', 'Author', 'Price', 'Pre-ordered Price', 'Sold']
ws.append(title)

header = {
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36'
}

# for loop in pages
for index in range(33):
    url = 'https://api.hahow.in/api/courses?limit=24&page='
    url = url + str(index)
    print(url)

    r = req.get(url, headers=header)
    print(r)

    root_json = r.json()
    # for loop in 24 course (in per page)
    for data in root_json['data']:
        course = []
        course.append(data['title'])  # course
        course.append(data['owner']['name'])  # author
        course.append(data['price'])  # price
        course.append(data['preOrderedPrice'])  # pre-ordered
        course.append(data['numSoldTickets'])  # sold
        ws.append(course)  # fill the list into the sheets

# save
wb.save('CourseData.xlsx')
