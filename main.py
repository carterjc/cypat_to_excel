import xlsxwriter
from bs4 import BeautifulSoup
import requests
import time

# TODO: Account for different divisions having a different number of images

imageMap = {}


def grabMainInfo(worksheet):
    main_page = requests.get("http://scoreboard.uscyberpatriot.org/")  # makes initial request, recieves HTML
    soup = BeautifulSoup(main_page.text, 'lxml')
    mainHTML = soup.select("tr")
    row = 1
    global imageNum
    imageNum = int(mainHTML[1].select("td")[5].text) * 2
    lastTeam = int(mainHTML[-1].select("td")[0].text)
    for team in mainHTML[1:]:
        rank = team.select("td")[0].text
        team_num = team.select("td")[1].text
        location = team.select("td")[2].text
        division = team.select("td")[3].text
        tier = team.select("td")[4].text
        play_time = team.select("td")[6].text
        warnings = team.select("td")[7].text
        image_score = team.select("td")[8].text
        adjustments = team.select("td")[9].text
        cisco_score = team.select("td")[10].text
        cumulative_score = team.select("td")[11].text
        worksheet.write(row, 0, rank)
        worksheet.write(row, 1, team_num)
        worksheet.write(row, 2, location)
        worksheet.write(row, 3, division)
        worksheet.write(row, 4, tier)
        worksheet.write(row, 5 + imageNum, play_time)
        worksheet.write(row, 6 + imageNum, warnings)
        worksheet.write(row, 7 + imageNum, image_score)
        worksheet.write(row, 8 + imageNum, adjustments)
        worksheet.write(row, 9 + imageNum, cisco_score)
        worksheet.write(row, 10 + imageNum, cumulative_score)
        row += 1
        print("Team " + str(rank) + "/" + str(lastTeam))
    print("Finished grabbing main data.")


def grabImageInfo(worksheet):
    main_page = requests.get("http://scoreboard.uscyberpatriot.org/")  # makes initial request, recieves HTML
    soup = BeautifulSoup(main_page.text, 'lxml')
    mainHTML = soup.select("tr")
    lastTeam = int(mainHTML[-1].select("td")[0].text)
    row = 1
    for team in mainHTML[1:]:
        rank = int(team.select("td")[0].text)
        team_page = requests.get("http://scoreboard.uscyberpatriot.org/team.php?team=" + team.select("td")[1].text)
        team_soup = BeautifulSoup(team_page.text, 'lxml')
        teamHTML = team_soup.select("tr")[3:]
        scores = []
        for image in teamHTML:
            scores.append((image.select("td")[0].text.split("_")[0], image.select("td")[5].text, image.select("td")[1].text))
            # Name, score, time
        for score in scores:
            worksheet.write(row, imageMap[score[0]], score[1])
            worksheet.write(row, imageMap[score[0]] + 1, score[2])
        row += 1
        time.sleep(2)
        print("Team " + str(rank) + "/" + str(lastTeam) + "{" + str(round((rank/lastTeam) * 100, 2)) + "%}")
    print("Finished grabbing image data.")


def createHeadings(worksheet):
    main_page = requests.get("http://scoreboard.uscyberpatriot.org/")  # makes initial request, recieves HTML
    soup = BeautifulSoup(main_page.text, 'lxml')
    mainHTML = soup.select("tr")
    images= []
    firstTeam = ""
    teamIndex = 1  # Index 0 is the table header
    while firstTeam == "":
        teamHTML = soup.select("tr")[1]
        if teamHTML.select("td")[3].text == "Open":  # Makes sure the team is of Open division
            firstTeam = teamHTML.attrs["href"][-7:]
            break
        teamIndex += 1
    # Makes specific team request
    team_page = requests.get("http://scoreboard.uscyberpatriot.org/team.php?team=" + firstTeam)
    soup = BeautifulSoup(team_page.text, 'lxml')
    numberOfImages = int(soup.select("tr")[1].findChildren()[4].text)  # Grabs the number of images
    image_pos = 5
    for i in range(3, numberOfImages + 3):
        fullName = soup.select("tr")[i].findChildren()[0].text
        image_name = fullName[:fullName.find("_")]
        images.append(image_name)  # Adds the refined image names to an array
        images.append(fullName[:fullName.find("_")] + " Time")
        global imageMap
        imageMap[image_name] = image_pos
        image_pos += 2
    worksheet.write(0, 0, "Rank")
    worksheet.write(0, 1, "Team Number")
    worksheet.write(0, 2, "Location")
    worksheet.write(0, 3, "Division")
    worksheet.write(0, 4, "Tier")
    worksheet.write(0, 5 + len(images), "Play Time")
    worksheet.write(0, 6 + len(images), "Warnings")
    worksheet.write(0, 7 + len(images), "Image Score")
    worksheet.write(0, 8 + len(images), "Adjustments")
    worksheet.write(0, 9 + len(images), "Cisco Score")
    worksheet.write(0, 10 + len(images), "Cumulative Score")
    for i in range(len(images)):
        worksheet.write(0, 5 + i, images[i])


def main():
    print("Welcome to the CyberPatriot scoreboard scraper!\nThis program will transform a current scoreboard "
          "into an excel file (including image scores and more)!\n")
    roundNumber = int(input("What round is it (number only)?"))
    season = str(input("What season is it (ex. CPXII)?"))
    workbook = xlsxwriter.Workbook(season + "r" + str(roundNumber) + ".xlsx")
    worksheet = workbook.add_worksheet()
    createHeadings(worksheet)
    grabMainInfo(worksheet)
    grabImageInfo(worksheet)
    workbook.close()
    print("Scores backed up.")


main()

# Col 0 = rank
# Col 1 = team num
# Col 2 = location
# Col 3 = division
# Col 4 = tier
# Col 5 = image score
# Col 6 = image score
# Col 7 = image score
# Col 8 = play time
# Col 9 = warnings
# Col 10 = image score
# Col 11 = Cisco score
# Col 12 = cumulative score
