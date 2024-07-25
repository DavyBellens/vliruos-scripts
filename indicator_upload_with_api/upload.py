import requests

url = "https://kuleuven.connect.vliruos.be/api/view/4698/import"
headers = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1HTHFqOThWTkxvWGFGZnBKQ0JwZ0I0SmFLcyIsImtpZCI6Ik1HTHFqOThWTkxvWGFGZnBKQ0JwZ0I0SmFLcyJ9.eyJhdWQiOiJhMjQ1OGMxYy1mMzRhLTRkMzEtYjRiOS02N2QwOWUwN2NkZjciLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC84YjljNjRkZC03MDkzLTQ5ZmUtOGM3Yy04OTEyYTk1NTc4YTQvIiwiaWF0IjoxNzIxODE3NzAxLCJuYmYiOjE3MjE4MTc3MDEsImV4cCI6MTcyMTgyMTYwMSwiYWlvIjoiQVdRQW0vOFhBQUFBN2FzUjhKaWZUSmthTVlPdzZSbWo2NlV1YXpmK3piSHQyd1pubU1ZOUJQQ2VKbUxabVdUNzJOdE5UdGtzaDBlNVNNS3BxckcvWExaU0pqVC81QnRIR240NDRwUmJhbDVFNStaYjNvQWZBSmJ0MmpDOVRzVkd2blV0KzhnQkVIRFoiLCJhbXIiOlsicHdkIl0sImVtYWlsIjoiZGF2eS5iZWxsZW5zQHZsaXJ1b3MuYmUiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9hZGI2NGZiNC02MmIzLTRlYmMtYmI5My0xNzM0N2I1NGFhODcvIiwiaXBhZGRyIjoiMmEwMjoyYzQwOjQwMDpiMDgwOjoxOmJkZjQiLCJuYW1lIjoiRGF2eSBCZWxsZW5zIiwibm9uY2UiOiJOMVl6VTJOS1VHWTBlazFhY0V4dVMzTkJkemN6WjFGdk5HOWZaRVp6V2psUE1HdDRjR2xaZWtsYVkxcEoiLCJvaWQiOiIwYjE4NDg5MC0yNTJkLTQ4MDgtODIxZi0zZTA1YjQ2ZTczNDMiLCJyaCI6IjAuQVU4QTNXU2NpNU53X2ttTWZJa1NxVlY0cEJ5TVJhSks4ekZOdExsbjBKNEh6ZmRQQUdjLiIsInJvbGVzIjpbIlN1cGVyQWRtaW4iXSwic3ViIjoiam9uVXhVdm5SN1RLeElmbXg0Zlk5dUtraXRac00zMmFCWUZFdldLSGI1ayIsInRpZCI6IjhiOWM2NGRkLTcwOTMtNDlmZS04YzdjLTg5MTJhOTU1NzhhNCIsInVuaXF1ZV9uYW1lIjoiZGF2eS5iZWxsZW5zQHZsaXJ1b3MuYmUiLCJ1dGkiOiI4aVZ0aFNiWmMwR2Jqc1BnY0FhR0FBIiwidmVyIjoiMS4wIn0.xb2DTmZVfNIpQ2Ylc_RWAeZRG2EewuM_f4wg8GqpSCaRTIE6H7THJ13-la4bzRdPr4mTKwpV8--9bmRrkh5r0Mg1K6-uCt6LyXOhpYDoKw1GYZ9ki9FQYRYOi2yBIWDBx7V-kI60AXghRzou4mJJLR1eQzat3EnLjXOWtIlVBOWC3nFHnhr31cbYNAGWd3MzySgrZdyr_DF3Ddpm7Ja2Gryohr9dI25N2QKtjbhmiticf1TxbfGjZwe08TMDUglDDwDSnogLh6EGoNNtRlQB0h10wb7_IX2kRoFca6kjOOkXDEC4ILx_St6DfIQouwCLSnpIpZTxZ30U1YKavNFwOg",
    "Connection": "keep-alive",
    "Content-Type": "multipart/form-data; boundary=----WebKitFormBoundary7N9N8IrNhIIM1otH",
    "Cookie": "ai_user=eJrk8xJ+9f34XPplnFajXu|2024-07-22T14:31:52.030Z; ARRAffinity=80c7a5f90186ed67b1a22b2d6ba64306a36f909c848093374b79f8b04109a233; ARRAffinitySameSite=80c7a5f90186ed67b1a22b2d6ba64306a36f909c848093374b79f8b04109a233; ai_session=QCPZXsFIBfMZfkiUl8tjG8|1721815791140|1721818001752",
    "Origin": "https://kuleuven.connect.vliruos.be",
    "Request-Id": "|4bbc7753b38b4be0b0bae3ee3aea81f3.e78765e9f5ad4024",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "Sec-GPC": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "sec-ch-ua": "\"Not/A)Brand\";v=\"8\", \"Chromium\";v=\"126\", \"Brave\";v=\"126\"",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "\"Windows\"",
    "traceparent": "00-4bbc7753b38b4be0b0bae3ee3aea81f3-e78765e9f5ad4024-01"
}


file_path = r"C:\Users\DavyBellens\VLIR-UOS\Digital Factory - APR-indicatoren 2022-2023\TEAM\TZ2022TEA530A101\7 - APR_Annex_1_TZ2022TEA530A101_Y1_VLIR_Final.xlsx"

files = {
    'dossierId': (None, '597'),
    file_path: (file_path, open(file_path, 'rb'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
    'overrideData': (None, 'true'),
    'moduleId': (None, '4698'),
    'viewUrl': (None, '.25883')
}

response = requests.post(url, headers=headers, files=files)

print(response.status_code)
print(response.text)
