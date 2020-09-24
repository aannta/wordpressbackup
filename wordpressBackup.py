#! python3
# Given URL to a public Wordpress blog's homepage
# Back up blog's content locally as Word documents
# starting page URL is in the format of either https://xxxxx.wordpress.com/
# or https://xxxxx.wordpress.com/01/01/2020/post-title-here/ if resuming previously interrupted back-up
# 'simply-formatted' Wordpress blogs are those with posts vertically arranged on main page
# and posts are stored in a chronological structure
# examples: https://audadulting.wordpress.com/ or https://omgtwinkles.wordpress.com/ 

import os, sys, time, random, re, requests, bs4, docx

# specify URL to Wordpress homepage
print('Enter Wordpress blog\'s homepage in the form of https://xxxxx.wordpress.com/')
homeURL = str(input()).replace(' ','')

# to determine whether element found is the innerpost p element
# @para bs4.element.tag
def isInnermostP(elem):
    strElem = str(elem)
    if strElem.count('<p>') == 1 and strElem.count('</p>') == 1:
        return True
    else:
        return False

# main program
try:
    
    # create local back-up folder
    if os.path.exists('WordpressBackup') is False:
        os.makedirs('WordpressBackup')

    # check if starting from actual homepage or a blogpost mid-way
    startPageRegex = re.compile(r'^https://(\w)+\.wordpress\.com/$')
    if startPageRegex.search(homeURL) is not None:
        print(f'Starting with homepage {homeURL}')
        # get URL to first post
        res = requests.get(homeURL)
        soup = bs4.BeautifulSoup(res.text, 'html.parser')
        elems = soup.select('a[href]')
        for i in range(len(elems)):
            link = elems[i].get('href')
            firstPageRegex = re.compile(re.escape(homeURL) + r'\d\d\d\d/\d\d/\d\d/')
            if firstPageRegex.search(link) is not None:
                firstPageURL = link
                break
        currPage = firstPageURL
    else:
        print(f'Starting with blogpost page {homeURL}')
        currPage = homeURL
    
    # starting trawling through the first page given
    while currPage is not None:
        
        # scrape current page
        print(f'Scraping page {currPage}...')
        res = requests.get(currPage)
        if res.status_code == 200:
            print('Sucessfully retrieved page\'s HTML')
        soup = bs4.BeautifulSoup(res.text, 'html.parser')

        # get current page's title
        if len(soup.select('.entry-title')) != 0:
            title = soup.select('.entry-title')[0].getText().replace('/', ' ').replace('*','').replace('?','')
        else:
            title = ' '
        
        # get current page's timestamp
        timeStamp = soup.select('time')[0].getText()
        timeStamp2 = soup.select('time')[0].get('datetime')

        # open new Word document, add post's title in bold and timestamp in italic
        d = docx.Document()
        d.add_paragraph(title)
        d.add_paragraph(timeStamp)
        d.paragraphs[0].runs[0].bold = True
        d.paragraphs[1].runs[0].italic = True
        
        # select either all innermost img elements, or all p elements
        postElems = soup.select('div[class=entry-content] img:empty, div[class=entry-content] p')
        
        for i in range(len(postElems)):
            
            # obtain image link either from data-orig-file (higher resolution) or src
            if (postElems[i].get('data-orig-file') is not None) or (postElems[i].get('src') is not None): 
                if postElems[i].get('data-orig-file') is not None:
                    imgRes = requests.get(postElems[i].get('data-orig-file'))
                else:
                    imgRes = requests.get(postElems[i].get('src'))
                imgFile = open(os.path.join('WordpressBackup', 'image.png'), 'wb')
                for chunk in imgRes.iter_content(100000):
                    imgFile.write(chunk)
                imgFile.close()
                d.add_picture(os.path.join('WordpressBackup', 'image.png'), width = docx.shared.Inches(6))
                os.unlink(os.path.join('WordpressBackup', 'image.png'))

            # for text element (only taking innermost element for pure text)
            elif len(postElems[i].getText()) > 0 and postElems[i].getText() != ' ' and isInnermostP(postElems[i]): 
                d.add_paragraph(postElems[i].getText())
        
        # save Word document with date and post title
        d.save(os.path.join('WordpressBackup', timeStamp2[:10] + ' ' + title + '.docx'))
        
        # find link to previous post
        prevLinkElem = soup.select('a[rel=prev]')
        if len(prevLinkElem) == 1:
            currPage = prevLinkElem[0].get('href')
        else:
            print('Final post of blog reached')
            break

        # stall randomly
        timeToSleep = random.randint(0, 10)
        print(f'Post successfully backed up. Sleeping for {timeToSleep} seconds...')
        time.sleep(timeToSleep)
        
    # print complete message
    print(f'Done backing up')

except Exception as err:
    print('An exception happened: ' + str(err))
