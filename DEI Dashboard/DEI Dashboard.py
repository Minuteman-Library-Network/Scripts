#!/usr/bin/env python3

"""
Jeremy Goldstein
Minuteman Library Network

Generates stats for DEI collections among our libraries
Updates data in a Google sheet used as a data source by Looker Studio
"""
# run in py313

import configparser
import psycopg2
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import traceback
import datetime
import pygsheets

# from oauth2client.service_account import ServiceAccountCredentials
# from googleapiclient.discovery import build


# function takes a sql query as a parameter, connects to a database and returns the results
def runquery(query):
    # read config file with database login details
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # Connecting to PostgreSQL database
    try:
        conn = psycopg2.connect(config["sql"]["connection_string"])
    except psycopg2.Error as e:
        print("Unable to connect to database: " + str(e))

    # Opening a session and querying the database
    cursor = conn.cursor()
    cursor.execute(query)
    # Storing the results in a variable. We'll use it later.
    rows = cursor.fetchall()
    # close database connection
    conn.close()
    # return variable containing query results
    return rows


# log items that were corrected to an existing Google Sheet
def appendToSheet(spreadSheetId, data):
    if not data:
        raise ValueError("Data must not be empty.")

    gc = pygsheets.authorize(
        service_file="C:\\Scripts\\Creds\\GSheet updater creds.json"
    )

    sh = gc.open_by_key(spreadSheetId)
    wks = sh.sheet1  # or sh.worksheet_by_title("My Sheet")

    # If the first value in the data is "Acton", clear the sheet preserving the header
    if data[0][0] == "Acton":
        header = wks.get_row(1)
        wks.clear()
        wks.update_row(1, header)
        first_empty_row = 2
    else:
        # Find the first empty row and insert data there
        first_empty_row = len(wks.get_all_values(include_tailing_empty_rows=False)) + 1
    rows_needed = first_empty_row + len(data) - 1

    # Expand the sheet if the data would exceed the current grid size
    if rows_needed > wks.rows:
        wks.add_rows(rows_needed - wks.rows)
    wks.update_values(f"A{first_empty_row}", data)


# converts psycopg2 fetchall() output to matrix required by pygsheets
def parse_pg_data(rows):

    def convert(val):
        if val is None:
            return ""
        if isinstance(val, (datetime.date, datetime.datetime)):
            return val.isoformat()  # e.g. "2026-03-04"
        return val  # int, float, str pass through as-is

    return [list(convert(val) for val in row) for row in rows]


# function constructs and sends outgoing email given a subject, a recipient and body text in both txt and html forms
def send_email_error(subject, message, recipient):
    # read config file with Sierra login credentials
    config = configparser.ConfigParser()
    config.read("C:\\Scripts\\Creds\\config.ini")

    # These are variables for the email that will be sent.
    # Make sure to use your own library's email server (emailhost)
    emailhost = config["email"]["host"]
    emailuser = config["email"]["user"]
    emailpass = config["email"]["pw"]
    emailport = config["email"]["port"]
    emailfrom = config["email"]["sender"]

    # Creating the email message
    msg = MIMEMultipart()
    emailmessage = message
    msg["From"] = emailfrom
    if type(recipient) is list:
        msg["To"] = ", ".join(recipient)
    else:
        msg["To"] = recipient
    msg["Date"] = formatdate(localtime=True)
    msg["Subject"] = subject
    msg.attach(MIMEText(emailmessage))

    # Sending the email message
    smtp = smtplib.SMTP(emailhost, emailport)
    # for Gmail connection used within Minuteman
    smtp.ehlo()
    smtp.starttls()
    smtp.login(emailuser, emailpass)
    smtp.sendmail(emailfrom, recipient, msg.as_string())
    smtp.quit()


def main(library, location):
    try:
        # enable creds file for referencing GSheets IDs
        config = configparser.ConfigParser()
        config.read("C:\\Scripts\\Creds\\config.ini")

        # query to gather dei collection report for each library location
        query = (
            """
    WITH topic_list AS (
      SELECT
        record_id,
        topic,
        is_fiction
        FROM (
          SELECT
            d.record_id,
            CASE
              WHEN REPLACE(d.index_entry,'.','') ~ '^\y(?!\w((ecology)|(ecotourism)|(ecosystems)|(environmentalism)|(african american)|(african diaspora)|(blues music)|(freedom trail)|(underground railroad)|(women)|(ethnic restaurants)|
(social life and customs)|(older people)|(people with disabilities)|(gay(s|\y(?!(head|john))))|(lesbian)|(bisexual)|(gender)|(sexual minorities)|(indian (art|trails))|(indians of)|(inca(s|n))|
(christian (art|antiquities|saints|shrine|travel))|(pilgrims and pilgrimages)|(jews)|(judaism)|((jewish|islamic) architecture)|(convents)|(sacred space)|(sepulchral monuments)|(spanish mission)|(spiritual retreat)|(temples)|(houses of prayer)|(religious institutions)|(monasteries)|(holocaust)|(church (architecture|buildings|decoration))))\w.*((guidebooks)|(description and travel))' THEN 'None of the Above'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(\yzen\y)|(dalai lama)|(buddhis)' THEN 'Buddhism'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(\yhindu(?!(stan|\skush)))|(divali)|(\yholi\y)|(bhagavadgita)|(upanishads)|(\ybrahman(s|ism))' THEN 'Hinduism'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(agnosticism)|(atheism)|(secularism)' THEN 'Agnosticism & Atheism'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(^\y(?!\w*terrorism)\w*(islam(?!.*(fundamentalism|terrorism))))|(\ysufi(sm)?)|(ramadan)|(id al (fitr\y)|(\yadha\y))|(quran)|(sunnites)|(shiah)|(muslim)|(mosques)|(qawwali)' THEN 'Islam'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(working class)|(social ((status)|(mobility)|(class)|(stratification)))|(standard of living)|(poor)|(\ycaste\y)|(classism)' THEN 'Class'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(south asia)|(indic\y)|(^\y(?!\w*k2)\w*(pakistan(?!.*k2)))|(\yindia\y)|(bengali)|(afghan(?!(\swar|s coverlets)))|(bangladesh)|(^\y(?!\w*everest)\w*(nepal(?!.*everest)))|(sri lanka)|(bhutan)|(east indian)' THEN 'South Asian'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(east asia)|(asian americans)|(^\y(?!\w*everest)\w*(chin(a(?!\sfictitious)|ese)(?!.*everest)))|(japan(?!ese beetle))|(korea(?!n war))|(taiwan)|(vietnam(?! war))|(cambodia)|(mongolia)|(lao(s|tian))|(myanmar)|(\ymalay)|((?<!muay )\ythai)|(philippin)|(indonesia)|(polynesia)|(brunei)|(east timor)|(pacific island)|(tibet autonomous)|(hmong)|(filipino)|(burm(a|ese(?! (python|cat))))' THEN 'East Asian & Pacific Islander'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(bullying)|(aggressiveness)|((?<!(substance|medication|opioid|oxycodone|cocaine|marijuana|opium|phetamine|drug|morphine|heroin))\sabuse(?!\sof administrative))|(violent crimes)|((?<!non)violence)|(crimes against)|((?<!(su)|(herb)|(pest))icide)|(suicide bomber)|(^\y(?!\w*investigation)\w*(murder(?!.*investigation)))|((human|child) trafficking)|(kidnapping)|(victims of)|(rape)|(police brutality)|(harassment)|(torture)' THEN 'Abuse & Violence'
	          WHEN REPLACE(d.index_entry,'.','') ~ '((?<!recordings for people.*)disabilit)|(blind)|(deaf)|(terminally ill)|(amputees)|(patients)|(aspergers)|(neurobehavioral)|(neuropsychology)|(neurodiversity)|(brain variation)|(personality disorder)|(autis(m|tic))|(barrier free design)' THEN 'Disabilities & Neurodiversity'
              WHEN REPLACE(d.index_entry,'.','') ~ '(acceptance)|(anxiety)|(compulsive)|(schizophrenia)|(eating disorders)|(mental(( health)|( illness)|( healing)|(ly ill)))|(resilience personality)|(suicid(?!e bomb))|(self (esteem|confidence|realization|perception|actualization|management|destructive|control))|(emotional problems)|(mindfulness)|(depressi(?!ons))|(stress (psychology|disorder))|(psychic trauma)|((?<!(homo|islamo|trans|xeno))phobia)' THEN 'Mental & Emotional Health'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(gamblers)|(drug use)|(alcoholi(?<!c beverages))|(addiction)|(drug use)|(substance|medication|opioid|oxycodone|cocaine|marijuana|opium|phetamine|drug|morphine|heroin)\sabuse|(binge drinking)|((?<!relationship )addict)' THEN 'Substance Abuse & Addiction'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(sexual minorities)|(gender)|(asexual)|(bisexual)|(gay(s|\y(?!(head|john))))|(intersex)|(homosexual)|(lesbian)|(stonewall riots)|(masculinity)|(femininity)|(trans(sex|phobia))|(drag show)|(male impersonator)|(queer)|(lgbtq)' THEN 'LGBTQIA+ & Gender Studies'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(indigenous)|(aboriginal)|((?<!east\s)\yindians(?!\sbaseball))|(trail of tears)|(aztecs)|(indian art)|(maya(s|n))|(eskimos)|(inuit)|(\yinca(s|n)\y)|(arctic peoples)|(aleut)|(american indian)|(indian reservations)|(maori)|(((abenaki)|(algonquian)|(apache)|(cherokee)|(chickasaw)|(chocktaw)|(cree)|(dakota)|(hopi)|(iroquois)|(kiowa)|(munduruku)|(navajo)|(ojibwa)|(oneida)|(osage)|(powhatan)|(pueblo)|(quiche)|(shoshoni)|(siksika)|(taino)|(tlingit)|(tuscarora)|(tzotzil)|(winnebago)|(yankton))\s((women)|(language)|(mythology)|(dance)|(silverwork)|(textile)|(nation)|(literature)|(long walk)))' THEN 'Indigenous'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(\yarab)|(middle east)|(palestin)|(bedouin)|((?<!(king of|putnam|potter|loring) )\yisrael(?!\slee))|(saudi)|(yemen)|(iraq(?!\swar))|(\yiran)|(\yegypt(?!ologists))|(leban(on|ese))|(qatar)|(syria)|((?<!wild )turk((ish|ey(?!(s| hunting)))?)\y)|(kurdis)|(bahrain)|(cyprus)|(kuwait)|(\yoman)|(?<!(belfort|lacey|romero|peele|kisner|lebowitz|miller|myles|reid|rubin|schnitzer|shakoor|sonnenblick|spieth|john|davis|clara|richard) )jordan(?!\s(ruth|fisher|vernon|michael|barbara|robbie|carol|john|david|grace|family|schnitzer|hal|louis|karl|raisa|dorothy|clarence|bruce|billy|andrew|b\y|wong|will|ted|steve|robert|pete|pat|mattie|marsh|leslie|june|joseph|hamilton|zach|teresa|bella|eben))' THEN 'Arab & Middle Eastern'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(hispanic)|((?<!new\s)(mexic))|(latin america)|(\ycuba(?!n\smissile))|(puerto ric)|(dominican)|(el salvador)|(salvadoran)|(argentin)|(bolivia)|
(chile)|(colombia)|(costa rica)|(ecuador)|(equatorial guinea)|(guatemala)|(hondura)|(nicaragua)|(panama)|(paragua)|(peru(?!gia))|(spain)|(spaniard)|(spanish)|(urugua)|(venezuela)|
((?<!jiu jitsu )brazil)|(guiana)|(guadeloup)|(martinique)|(saint barthelemy)|(saint martin)' THEN 'Hispanic & Latino'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(\yafro)|(blacks(?!mith))|(men black)|((?<!game reserve south )africa(?!(nized|nus|n (literature french|elephant|gray|black rhino|buffalo|wild dog|python|pygmy|violets))))|(black (nationalism|panther party|power|muslim|lives))|(harlem renaissance)|(abolition)|(segregation)|(^\y(?!\w*((rome)|(italy)|(egypt)))\w*(slave(s|(ry)?)(?!((rome)|(egypt)|(italy)))))|(emancipation)|(underground railroad)|(apartheid)|((?<!kincaid )jamaica)|(haiti)|(nigeria)|((?<!cepheus king of )ethiopia)|(^\y(?!\w*(bonobo|wildlife))\w*congo)|(^\y(?!\w*kilmanjaro)\w*(tanzania(?!.*kilmanjaro)))|((?<!(mammals|elephants|leopard|cows|lion|animals|zoology|conservation|behavior) )kenya)|(uganda)|(sudan)|(ghana)|(cameroon)|
((?<!(conservation|animals|jungles|species|lemurs|dinosaurs|fossil|zoology|wildlife watching) )madagascar)|(mozambique)|(angola)|(cote divoire)|(\ymali\y)|(burkina faso)|(malawi)|(somalia)|(zambia)|(senegal)|(zimbabw)|((?<!gorilla )rwanda)|
(eritrea)|(guinea (?!pig))|(benin\y)|(burundi)|(sierra leone)|(\ytogo\y(?! dog))|(liberia)|(mauritania)|(\ygabon)|(namibia)|
(botswana)|(lesotho)|(gambia)|(eswatini)|(djibouti)|(\ytutsi\y)|((?<!(daybell|johnson|foster|gardenier|gibbs|hurley|jenkins|kerley|kister|rje) )\ychad\y)' THEN 'Black'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(jewish)|(jews)|(judaism)|(hanukkah)|(purim)|(passover)|(zionis)|(hasidism)|(antisemitism)|(rosh hashanah)|(yom kippur)|(sabbath)|(sukkot)|(pentateuch)|(synagogue)|(hebrew)|(yiddish)|(seder)|(cabala)' THEN 'Judaism'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(genocide)|(equality)|(immigra)|(feminis)|(womens rights)|(sexism)|((?<!(fugitives from |young |jones ))justice(?!(s of the peace)|(\s(league|society|donald|benjamin|victoria))))|(racism)|(suffrag)|(sex role)|(social ((change)|(movements)|(problems)|(reformers)|(responsibilit)|(conditions)))|(sustainable development)|(environmental)|(poverty)|(abortion)|((human|civil) rights)|(prejudice)|(protest movements)|(homeless)|(public (health|welfare))|(discrimination)|(refugee)|((anti nazi|pro choice|labor) movement)|(race awareness)|(political prisoner)|(ku klux klan)|(colorism)|(activis)|(persecution)|(xenophobia)|(((privilege)|(belonging)|(alienation)|(stigma)|(stereotypes)) social)|(noncitizen)|(stateless person)|(deportation)|(abuse of power)|(boat people)' THEN 'Equity & Social Issues'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(multicultural)|(cross cultural)|(diasporas)|((?<!sexual )minorities)|(interracial)|(ethnic identity)|((race|ethnic) relations)|(racially mixed)|(bilingual)|(passing identity)' THEN 'Multicultural'
	          WHEN REPLACE(d.index_entry,'.','') ~ '(protestant)|(bible)|(nativity)|(adventis)|(mormon)|(baptist)|(catholic)|(methodis)|(pentecost)|(episcopal)|(lutheran)|(clergy)|((?<!(christ|mary\s|ezra\s))church(?!(ill|\sbenjamin|\sr w richard|\sfrederic|\sf forrester)))|(evangelicalism)|((?<!(siriano|amanpour|dior) )christian(?!(sen|son| dior))(?!.*\d{4}))|(easter\y)|(christmas)|(shaker)|(noahs ark)|(biblical)|(new testament)' THEN 'Christianity'
	          ELSE 'None of the Above'
            END AS topic,
            CASE
              WHEN d.index_entry ~ '((\yfiction)|(pictorial works)|(tales)|(^\y(?!\w*biography)\w*(comic books strips etc))|(^\y(?!\w*biography)\w*(graphic novels))|(\ydrama)|((?<!hi)stories))(( [a-z]+)?)(( translations into [a-z]+)?)$' AND b.material_code NOT IN ('7','8','b','e','j','k','m','n')
                AND NOT (ml.bib_level_code = 'm' AND ml.record_type_code = 'a' AND f.p33 IN ('0','e','i','p','s','','c')) THEN TRUE
              ELSE FALSE
            END AS is_fiction	

          FROM sierra_view.bib_record_location bl
          LEFT JOIN sierra_view.phrase_entry d
            ON bl.bib_record_id = d.record_id
            AND d.index_tag = 'd'
            AND d.is_permuted = FALSE
          JOIN sierra_view.bib_record_property b
            ON bl.bib_record_id = b.bib_record_id
          LEFT JOIN sierra_view.control_field f
            ON b.bib_record_id = f.record_id
          LEFT JOIN sierra_view.leader_field ml
            ON b.bib_record_id = ml.record_id

          WHERE bl.location_code ~ '^"""
            + location
            + """'
        )inner_query

      GROUP BY 1,2,3
    )

    SELECT *

    FROM (
      SELECT 
        '"""
            + library
            + """' AS library,
        mat.name AS format,
        t.topic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'j' AND t.is_fiction IS TRUE) AS juv_fic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'j' AND t.is_fiction IS FALSE) AS juv_nonfic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'y' AND t.is_fiction IS TRUE) AS ya_fic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'y' AND t.is_fiction IS FALSE) AS ya_nonfic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) NOT IN('y','j') AND t.is_fiction IS TRUE) AS adult_fic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) NOT IN('y','j') AND t.is_fiction IS FALSE) AS adult_nonfic,
        COUNT(DISTINCT i.id) AS total_items

      FROM sierra_view.item_record i
      JOIN sierra_view.bib_record_item_record_link l
        ON i.id = l.item_record_id
        AND i.location_code ~ '^"""
            + location
            + """'
      JOIN topic_list t
        ON l.bib_record_id= t.record_id
        AND t.topic != 'None of the Above'
      JOIN sierra_view.bib_record_property b
        ON t.record_id = b.bib_record_id
      JOIN sierra_view.material_property_myuser mat
        ON b.material_code = mat.code

      GROUP BY 1,2,3

      UNION

      SELECT 
        '"""
            + library
            + """' AS library,
        mat.name AS format,
        'Unique Diverse Items' AS topic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'j' AND t.is_fiction IS TRUE) AS juv_fic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'j' AND t.is_fiction IS FALSE) AS juv_nonfic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'y' AND t.is_fiction IS TRUE) AS ya_fic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'y' AND t.is_fiction IS FALSE) AS ya_nonfic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) NOT IN('y','j') AND t.is_fiction IS TRUE) AS adult_fic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) NOT IN('y','j') AND t.is_fiction IS FALSE) AS adult_nonfic,
        COUNT(DISTINCT i.id) AS total_items

      FROM sierra_view.item_record i
      JOIN sierra_view.bib_record_item_record_link l
        ON i.id = l.item_record_id 
        AND i.location_code ~ '^"""
            + location
            + """'
      JOIN topic_list t
        ON l.bib_record_id= t.record_id
        AND t.topic != 'None of the Above'
      JOIN sierra_view.bib_record_property b
        ON t.record_id = b.bib_record_id
      JOIN sierra_view.material_property_myuser mat
        ON b.material_code = mat.code

      GROUP BY 1,2,3

      UNION

      SELECT
        '"""
            + library
            + """' AS library,
        mat.name AS format,
        'None of the Above' AS topic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'j' AND t.is_fiction IS TRUE) AS juv_fic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'j' AND t.is_fiction IS FALSE) AS juv_nonfic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'y' AND t.is_fiction IS TRUE) AS ya_fic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) = 'y' AND t.is_fiction IS FALSE) AS ya_nonfic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) NOT IN('y','j') AND t.is_fiction IS TRUE) AS adult_fic,
        COUNT(DISTINCT i.id) FILTER(WHERE SUBSTRING(i.location_code,4,1) NOT IN('y','j') AND t.is_fiction IS FALSE) AS adult_nonfic,
        COUNT(DISTINCT i.id) AS total_items

      FROM sierra_view.item_record i
      JOIN sierra_view.bib_record_item_record_link l
        ON i.id = l.item_record_id
        AND i.location_code ~ '^"""
            + location
            + """'
      JOIN (
        SELECT
          t.record_id,
          t.is_fiction
        FROM topic_list t
        GROUP BY 1,2
        HAVING COUNT(DISTINCT t.topic) FILTER (WHERE t.topic != 'None of the Above') = 0
       ) t
         ON l.bib_record_id= t.record_id
      JOIN sierra_view.bib_record_property b
        ON t.record_id = b.bib_record_id
      JOIN sierra_view.material_property_myuser mat
      ON b.material_code = mat.code

      GROUP BY 1,2,3
    )a

    ORDER BY 1,2,
    CASE
      WHEN topic = 'Unique Diverse Items' THEN 2
      WHEN topic = 'None of the Above' THEN 3
      ELSE 1
    END, topic
    """
        )

        results = runquery(query)
        parsed_results = parse_pg_data(results)
        appendToSheet(config["gsheet"]["dei"], parsed_results)
    except Exception:
        # read config file with recipient list for email
        config_recipient = configparser.ConfigParser()
        config_recipient.read("C:\\Scripts\\Creds\\emails.ini")
        emailto = config_recipient["script_error"]["recipients"].split()

        # craft email subject and message containing error message details from traceback
        email_subject = "DEI Dashboard " + library + " script error"
        email_message = (
            "Your script failed with the following error:\n\n" + traceback.format_exc()
        )

        send_email_error(email_subject, email_message, emailto)
        raise


if __name__ == "__main__":
    main("Acton", "^act")
    main("Acton/West", "^ac2")
    main("Arlington", "^arl")
    main("Arlington/Fox", "^ar2")
    main("Ashland", "^ash")
    main("Bedford", "^bed")
    main("Belmont", "^blm")
    main("Brookline", "^brk")
    main("Brookline/Coolidge Corner", "^br2")
    main("Brookline/Putterham", "^br3")
    main("Cambridge", "^cam")
    main("Cambridge/Outreach", "^ca3")
    main("Cambridge/Boudreau", "^ca4")
    main("Cambridge/Central Square", "^ca5")
    main("Cambridge/Collins", "^ca6")
    main("Cambridge/O" "Connell", "^ca7")
    main("Cambridge/O" "Neill", "^ca8")
    main("Cambridge/Valente", "^ca9")
    main("Concord", "^con")
    main("Concord/Fowler", "^co2")
    main("Dedham", "^ddm")
    main("Dedham/Endicott", "^dd2")
    main("Dean", "^dea")
    main("Dover", "^dov")
    main("Framingham Public", "^fpl")
    main("Framingham Public/McAuliffe", "^fp2")
    main("Framingham State", "^fst")
    main("Franklin", "^frk")
    main("Holliston", "^hol")
    main("Lasell", "^las")
    main("Lexington", "^lex")
    main("Lincoln", "^lin")
    main("Maynard", "^may")
    main("Medfield", "^mld")
    main("Medford", "^med")
    main("Medway", "^mwy")
    main("Millis", "^mil")
    main("Natick", "^na(t|4)")
    main("Natick", "^na2")
    main("Needham", "^nee")
    main("Newton", "^ntn")
    main("Norwood", "^nor")
    main("Olin", "^oln")
    main("Regis", "^reg")
    main("Sherborn", "^shr")
    main("Somerville", "^som")
    main("Somerville/East", "^so2")
    main("Somerville/West", "^so3")
    main("Stow", "^sto")
    main("Sudbury", "^sud")
    main("Waltham", "^wlm")
    main("Watertown", "^wat")
    main("Wayland", "^wyl")
    main("Wellesley", "^wel")
    main("Wellesley", "^we2")
    main("Wellesley", "^we3")
    main("Weston", "^wsn")
    main("Westwood", "^wwd")
    main("Westwood/Islington", "^ww2")
    main("Winchester", "^win")
    main("Woburn", "^wob")
