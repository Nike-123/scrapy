import csv, json
from pathlib import Path

src = Path('/home/user/scrapy/ip_verification/batch2/output/verified_ip_firms_batch2.csv')
rows = list(csv.DictReader(src.open(encoding='utf-8', newline='')))
assert len(rows) == 60 and [int(r['Row_ID']) for r in rows] == list(range(51, 111))

ENTITY = {
 51:'Law firm',52:'Engineering consultancy',53:'Unknown / inactive',54:'Law firm',
 55:'Unknown / inactive',56:'Law firm',57:'Law firm',58:'IP advisory / asset-management firm',
 59:'Unknown / inactive',60:'Unknown / inactive',61:'Solo law practice',62:'Unknown / inactive',
 63:'Law firm',64:'Unknown / inactive',65:'Law firm',66:'Law firm',
 67:'Individual / academic (non-legal)',68:'Law firm',69:'Patent brokerage (non-law-firm)',
 70:'Patent search firm (non-law-firm)',71:'Solo law practice',72:'Unknown / inactive',
 73:'Law firm',74:'Solo law practice',75:'Law firm',76:'Law firm',77:'Law firm',
 78:'Solo law practice',79:'Solo law practice',80:'Law firm',81:'Law firm',82:'Law firm',
 83:'Law firm',84:'Patent research firm (non-law-firm)',85:'Unknown / inactive',
 86:'Law firm',87:'Solo law practice',88:'Unknown / inactive',
 89:'Patent drafting vendor (non-law-firm)',90:'Unknown / inactive',91:'Unknown / inactive',
 92:'Law firm',93:'Law firm',94:'Registered patent-agent practice',95:'Law firm',
 96:'Unknown / inactive',97:'Legal-tech vendor (non-law-firm)',98:'Solo law practice',
 99:'Law firm (merged into Scale LLP)',100:'IP advisory / investment firm (non-law-firm)',
 101:'Law firm',102:'Unknown / inactive',103:'Solo law practice',104:'Solo law practice',
 105:'Solo law practice',106:'Solo law practice',107:'IP filing/prosecution services firm',
 108:'Unknown / inactive',109:'Law firm',110:'Law-firm marketing consultancy (non-law-firm)',
}
# Owner_Evidence_URL: only where that page actually names the person
OWNER_EV = {
 51:'https://dvoraklawgroup.com/dvorak-law-group-to-expand-headquarters/',
 52:'https://www.dwgildart.com/',57:'https://www.el-law.com/',
 61:'https://www.eddingtonlawoffice.com/',63:'https://eichiplaw.com/',
 65:'https://uslawpros.com/',66:'https://www.eleylawfirm.com/',
 67:'https://elizabethchurchill.com/',68:'https://www.kimassociates.com/',
 70:'https://www.epsilonpatents.com/',71:'https://www.ergenzingeriplaw.com/',
 74:'https://esquiretrademarks.com/',75:'https://essentialpatent.net/david-w-long/',
 76:'https://estriatuslaw.com/',77:'https://etheridgelaw.com/',
 78:'https://www.ethertonlaw.com/',79:'https://www.evaguo.com/',
 82:'https://www.evermanlawfirm.com/team.html',
 86:'https://www.fabianolawfirm.com/aboutflfpc.html',
 87:'https://www.falkowskipllc.com/CJF%20BIO.htm',
 94:'https://www.farsightpatent.com/',98:'https://fasthofflawfirm.com/',
 101:'https://www.fattibene.com/pages/812622/index.htm',
 103:'https://www.fehrlawfirm.com/attorneybio.html',
 104:'https://feinberglaw.com/',105:'https://www.ferguslegalip.com/',
 106:'https://ferrancepc.com/',109:'https://www.finniplaw.com/',
}
EMAIL_STATUS = {
 79:'Published on official website (Gmail address bearing owner name)',
 86:'Published professional email on official contact page',
 87:'Published professional email on official attorney bio page',
 101:'Published professional email in official website contact block',
 109:'Published professional email on official website (founder section)',
}
GEN_SRC_OVERRIDE = {58:'https://www.echelonip.com/Contact.html'}

def source_url(orig):
    o = orig.strip()
    if not o.lower().startswith(('http://','https://')):
        o = 'https://' + o
    return o

mapped, source_rows, review_rows, shortlist = [], [], [], []
for r in rows:
    rid = int(r['Row_ID'])
    gen_src = ''
    if r['General_Email']:
        gen_src = GEN_SRC_OVERRIDE.get(rid, r['Verified_Official_Website'])
    m = [
        r['Row_ID'], r['Company_Name_Original'], r['Website_URL_Original'],
        r['Company_Name_Verified'], ENTITY[rid], r['Website_Status'],
        r['Verified_Official_Website'], r['US_Based_Status'], r['US_Evidence_URL'],
        r['IP_Status'], r['IP_Practice_Areas'], r['IP_Evidence_URL'],
        r['Owner_First_Name'], r['Owner_Last_Name'], r['Owner_Title'],
        OWNER_EV.get(rid, ''), r['Owner_Email'], r['Email_Source_URL'],
        EMAIL_STATUS.get(rid, ''), r['General_Email'], gen_src,
        r['Verification_Status'], r['Review_Notes'], r['Checked_Date'],
    ]
    assert len(m) == 24, rid
    if r['Owner_Email']:
        assert rid in EMAIL_STATUS, ('missing email status', rid)
    mapped.append(m)
    source_rows.append([r['Row_ID'], r['Company_Name_Original'], r['Website_URL_Original'], source_url(r['Website_URL_Original'])])
    if r['Verification_Status'] == 'Needs Review':
        review_rows.append([r['Row_ID'], r['Website_URL_Original'], r['Company_Name_Verified'], r['Verification_Status'], r['IP_Status'], r['US_Based_Status'], r['Review_Notes']])
    if r['US_Based_Status'] == 'US-Based' and r['IP_Status'] == 'IP':
        shortlist.append(m)

print('mapped rows:', len(mapped), '| review queue adds:', len(review_rows), '| shortlist:', len(shortlist))
print('shortlist Row_IDs:', [m[0] for m in shortlist])
print('review Row_IDs:', [x[0] for x in review_rows])
print('owner-email rows:', [r['Row_ID'] for r in rows if r['Owner_Email']])
print('general-email rows:', [r['Row_ID'] for r in rows if r['General_Email']])

def dump(name, data, chunk):
    for i in range(0, len(data), chunk):
        part = data[i:i+chunk]
        Path(f'/tmp/gs_{name}_{i//chunk+1}.json').write_text(json.dumps(part, ensure_ascii=False), encoding='utf-8')
        print(f'/tmp/gs_{name}_{i//chunk+1}.json', len(part), 'rows')

dump('verification', mapped, 20)
dump('source', source_rows, 30)
dump('review', review_rows, 37)
dump('shortlist', shortlist, 22)
