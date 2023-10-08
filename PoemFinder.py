import requests, bs4, openpyxl

base_url = "https://www.poemhunter.com/poems/love/page-{}/"

# Load the existing workbook and get the first worksheet
filepath = "C:\\Users\\mamma\\OneDrive\\Desktop\\poemFinder.xlsx"
workbook = openpyxl.load_workbook(filepath)
worksheet = workbook.active

# Find the last filled row in the worksheet
row_number = worksheet.max_row + 1

for page_num in range(1, 99):
    try:
        result = requests.get(base_url.format(page_num))
        soup = bs4.BeautifulSoup(result.text, "lxml")

        links = soup.select(".phLink")
        for link in links[2:]:
            poem_url = link.find("a")["href"]
            print(poem_url)
            
            newResult = requests.get("https://www.poemhunter.com" + poem_url)
            newSoup = bs4.BeautifulSoup(newResult.text, "lxml")
            test = newSoup.select(".phContainer")[2].select(".phcText")[0]

            poem_text = '\n'.join([br.next_sibling.strip() if br.next_sibling else " " for br in test.find_all("br")])
            output = poem_text.encode('utf-8', errors='ignore').decode('cp1252', errors='ignore')
            
            # Append the poem to the worksheet
            worksheet.cell(row=row_number, column=1, value=output)
            row_number += 1

    except Exception as e:
        print(f"An error occurred: {e}")

# Save the workbook
workbook.save(filepath)