""" ---------------------------------------------------------------------------
    Description
    --------------------------------------------------------------------------- 

    Provide several statistics from Outlook.
    Generate an HTML report
    (Heavily co-authored by ChatGPT)
    
    Horrible programing. This could be way more consistent 
    than it is right now, especially on date management.
 
    To add: top meeting inviters I don't answer to.
    To add: top senders I don't read
    To add: Top senders I delete
    To add: Top senders putting me in BCC
 
    Author:         JL Dupont
    Date/version:   20230211
"""

""" ---------------------------------------------------------------------------
    Imports
    ----------------------------------------------------------------------- """
import  pytz
import  win32com.client
import  webbrowser
import  plotly.graph_objects    as      go
from    collections             import  Counter
from    datetime                import  timedelta, datetime

""" ---------------------------------------------------------------------------
    Constants
    ----------------------------------------------------------------------- """
C_BANNEDFOLDERS             = [     'PersonMetag_dir', 
                                    'Social Activity Notifications',
                                    'Quick Step Settings',
                                    'RSS Subscriptions',
                                    'Sync Issues',
                                    'Conversation Action Settings',
                                    'Calendar',
                                    'Conversation History',
                                    'Contacts',
                                    'ExternalContacts',
                                    'Journal',
                                    'PersonMetadata',
                                    'Notes',
                                    'Tasks',
                                    'Yammer Root',
                                    'Files' ]
C_REPORTS                   = {     'readunread':       ['01.html', 'Read/Unread emails by folder'],
                                    'topsenders':       ['02.html', 'Top senders'],
                                    'toprecipients':    ['03.html', 'Top recipients (People I sent emails to)'],
                                    'topccsenders':     ['04.html', 'Top senders who put me in copy'],
                                    'topgroupsenders':  ['05.html', 'Top senders who put me in a mailing list'],
                                    'topinviters':      ['06.html', 'Top people who sents me meeting invitations'],
                                    'meetinganswers':   ['07.html', 'How I answered to meeting invitations'],
                                    'longestthread':    ['',        'The longest thread in my inbox'],
                                    'meetings':         ['',        'More information about meetings'] }
C_VERSION                   = "20230211"
C_TOPLENGTH                 = 5
C_NBDAYS                    = 365
C_HTMLFILE                  = 'outlookstats.html'
C_INFOEND                   = '<p></p></article></div></div>'
C_INFOSTART                 = '<div class="row"><div class="col-2"><article class="post">'
C_GREEN                     = '#20B2AA'
C_RED                       = '#DB7093'
C_BLUE                      = '#6A5ACD'

""" ---------------------------------------------------------------------------
    Global variables
    ----------------------------------------------------------------------- """
g_dir                       = []
g_lastyear                  = datetime.now(pytz.utc) - timedelta(days=C_NBDAYS)
g_lastyearfilter            = " AND [ReceivedTime] >= '" + g_lastyear.strftime('%m/%d/%Y %H:%M %p') + "'"
g_tz                        = pytz.timezone('UTC')


""" ---------------------------------------------------------------------------
    Functions
    ----------------------------------------------------------------------- """
    
def f_longestthread(folder):
    '''
    List the top X longest thread of the year
    
    Parameters:
        folder (object)     : The folder object which contains the emails.

    Return:
        count (int)         : The length of the longest thread
        subject (str)       : The subject of the longest thread
    '''
    
    threads                 = {}
    for mail in folder.Items:
        if mail.CreationTime > g_lastyear:
            thread_id       = mail.ConversationID
            if thread_id in threads:
                threads[thread_id]['count'] += 1
                if mail.CreationTime > threads[thread_id]['latest_time']:
                    threads[thread_id]['latest_time'] = mail.CreationTime
                threads[thread_id]['subjects'].append(mail.Subject)
            else:
                threads[thread_id] = {'count': 1, 'latest_time': mail.CreationTime, 'subjects':[mail.Subject]}
    longest_threads = sorted(threads.items(), key=lambda x: x[1]['count'], reverse=True)[C_TOPLENGTH]
    count                   = [thread[1]['count'] for thread in longest_threads]
    subject                 = [thread[1]['subjects'][0] for thread in longest_threads]
    return count, subject


def f_readunread(folder):
    '''
    Retrieve the count of unread and read emails in a folder and its subfolders.
    The function is reccursive.
    
    Parameters:
        folder (object)     : The Outlook folder object which contains the emails.
    
    Return:
        No return value. Use g_dir as a global variable to store results
    '''
    
    if folder.Name in C_BANNEDFOLDERS:
        return
    if folder.Items.Count > 0:
        unread              = folder.Items.Restrict("[UnRead] = true"   + g_lastyearfilter).Count
        read                = folder.Items.Restrict("[UnRead] = false"  + g_lastyearfilter).Count
        g_dir.append([folder.Name,read,unread])
    for subfolder in folder.Folders:
        f_readunread(subfolder)


def f_topsenders(folder):
    '''
    Return the top N senders in an outlook folder
    
    Parameters:
        folder (object)     : The Outlook folder object which contains the emails.
    
    Return:
        A list of tuples    : {sendername(str),count(int)}
    '''
    
    senderlist              = []
    for msg in folder.Items:
        try: 
            if msg.SenderName != "Microsoft Outlook":
                creation_time = msg.CreationTime.astimezone(g_tz)
                if (datetime.now(g_tz) - creation_time).days <= C_NBDAYS:
                    senderlist.append(msg.SenderName)
        except:
            pass    
    count                   = Counter(senderlist)
    return sorted(count.items(), key=lambda x: x[1], reverse=True)[:C_TOPLENGTH]


def f_topccsenders(folder):
    '''
    Return the top N senders in an outlook folder where I am in CC
    
    Parameters:
        folder (object)     : The Outlook folder object which contains the emails.
    
    Return:
        A list of tuples    : {sendername(str),count(int)}
    '''
    
    cc_senderlist           = []
    for msg in folder.Items:
        try: 
            if msg.SenderName != "Microsoft Outlook":
                creation_time = msg.CreationTime.astimezone(g_tz)
                if (datetime.now(g_tz) - creation_time).days <= C_NBDAYS:
                    for recipient in msg.Recipients:
                        if recipient.Type == 2 and recipient.Address == outlook.Session.CurrentUser.Address:
                            cc_senderlist.append(msg.SenderName)
        except:
            pass
    count                   = Counter(cc_senderlist)
    return sorted(count.items(), key=lambda x: x[1], reverse=True)[:C_TOPLENGTH]


def f_topgroupsenders(folder):
    '''
    Return the top N senders in an outlook folder where I am in mailing lists
    
    Parameters:
        folder (object)     : The Outlook folder object which contains the emails.
    
    Return:
        A list of tuples    : {sendername(str),count(int)}
    '''

    senderlist              = []
    for msg in folder.Items:
        try: 
            if msg.SenderName != "Microsoft Outlook":
                creation_time = msg.CreationTime.astimezone(g_tz)
                if (datetime.now(g_tz) - creation_time).days <= C_NBDAYS:
                    recipient_list = []
                    for recipient in msg.Recipients:
                        recipient_list.append(recipient.Address)
                    if outlook.Session.CurrentUser.Address not in recipient_list:
                        senderlist.append(msg.SenderName)
        except Exception as e:
            pass
    count                   = Counter(senderlist)
    return sorted(count.items(), key=lambda x: x[1], reverse=True)[:C_TOPLENGTH]


def f_toprecipients(folder):
    '''
    Return the top recipient in an outlook folder 
    
    Parameters:
        folder (object)     : The Outlook folder object which contains the emails.
    
    Return:
        A list of tuples    : {sendername(str),count(int)}
    '''
    
    senderlist              = []
    for msg in folder.Items:
        try:
            reciplist = msg.Recipients
            creation_time = msg.CreationTime.astimezone(g_tz)
            if (datetime.now(g_tz) - creation_time).days <= C_NBDAYS:
                for recip in reciplist:
                    senderlist.append(str(recip.AddressEntry)) 
        except:
            pass
    count                   = Counter(senderlist)            
    return sorted(count.items(), key=lambda x: x[1], reverse=True)[:C_TOPLENGTH]


def f_meetingtime(folder):
    '''
    Return the time spent in meetings 
    
    Parameters:
        folder (object)     : The Outlook folder object (calendar expected) which contains the meeting.
    
    Return:
        totaltime (int)     : time spent in meetings
        totaltimebyme (int) : of which time spent in meetings that I created
    '''

    totaltime               = timedelta()
    totaltimebyme           = timedelta()
    appointments            = []
    for appointment in folder.Items:
        start   = appointment.Start.astimezone(pytz.utc)
        end     = appointment.End.astimezone(pytz.utc)
        if start > g_lastyear and appointment.MeetingStatus != 7:
            appointments.append((start, end, appointment.Subject, appointment.Organizer, appointment.MeetingStatus))
    appointments.sort(key=lambda x: x[0])
    for i in range(len(appointments)):
        start1, end1, subject, organizer, status = appointments[i]
        j       = i + 1
        while j < len(appointments) and appointments[j][0] < end1:
            start2, end2, _, _, status = appointments[j]
            if end1 > start2:
                end1 = min(end1, end2)
            j += 1
        if (end1-start1).total_seconds() < 8*3600:
            totaltime += end1 - start1
            #if organizer == outlook.Session.CurrentUser.Address:
            if status == 0:
                totaltimebyme += end1 - start1
    return int(totaltime.total_seconds()/3600), int(totaltimebyme.total_seconds()/3600)


def f_topmeetinginviters(calendar):
    appointments = calendar.Items
    appointments.Sort("[Start]")
    appointments = appointments.Restrict("[Start] >= '" + g_lastyear.strftime('%m/%d/%Y %H:%M %p') + "'")
    organizers = []
    appointment = appointments.GetFirst()
    while appointment:
        organizers.append(appointment.Organizer)
        appointment = appointments.GetNext()
    organizer_counts = Counter(organizers)
    top_organizers = dict(organizer_counts.most_common(5))
    return top_organizers


def f_conflictingmeetings(calendar):
    appointments = calendar.Items
    appointments.Sort("[Start]")
    appointments = appointments.Restrict("[Start] >= '" + g_lastyear.strftime('%m/%d/%Y %H:%M %p') + "'")
    start_times = []
    appointment = appointments.GetFirst()
    while appointment:
        start_times.append(appointment.Start)
        appointment = appointments.GetNext()
    start_time_counts = Counter(start_times)
    simultaneous_meetings = {start_time: count for start_time, count in start_time_counts.items() if count > 1}
    return len(simultaneous_meetings)


def f_meetinganswers(calendar):
    appointments = calendar.Items
    appointments.Sort("[Start]")
    appointments = appointments.Restrict("[Start] >= '" + g_lastyear.strftime('%m/%d/%Y %H:%M %p') + "'")
    accepted_count = 0
    noresponse_count = 0
    appointment = appointments.GetFirst()
    while appointment:
        #print(str(appointment.Organizer) + str(appointment.CreationTime.astimezone(g_tz)) +str(appointment.ResponseStatus) + str(appointment.Subject))
        if appointment.ResponseStatus == 3:
            accepted_count += 1
        if appointment.ResponseStatus == 5:
            noresponse_count += 1
        appointment = appointments.GetNext()
    return [('Accepted',accepted_count), ('Did not answer',noresponse_count)]

def f_htmlsection(filehandle, reporlist):
    with open(reporlist[0], "r") as input_file:
        filehandle.write(C_INFOSTART + input_file.read() + '<h1>' + reporlist[1] + '</h1>'  + C_INFOEND)

""" ---------------------------------------------------------------------------
    Main
    ----------------------------------------------------------------------- """

#-- Connect to outlook
print("""
 __       ___       __   __        __  ___      ___  __  
/  \ |  |  |  |    /  \ /  \ |__/ /__`  |   /\   |  /__` 
\__/ \__/  |  |___ \__/ \__/ |  \ .__/  |  /~~\  |  .__/                                                       
""")
print("OutlookStats (ver: " + C_VERSION + ")")

outlook                     = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox                       = outlook.GetDefaultFolder(6)
calendar                    = outlook.GetDefaultFolder(9)
sent                        = outlook.GetDefaultFolder(5)
me                          = str(outlook.Session.CurrentUser)
print("\nGenerating stats for: " + str(me))
print("Be patient!\n\n")

#-- number of emails read/unread by folder
print("Counting the number of read/unread emails by folder...")
root_folder                 = outlook.Folders
for folder in root_folder:
    f_readunread(folder)
g_dir                       = sorted(g_dir, key=lambda x: x[1], reverse=True)
fig = go.Figure(data=[go.Bar(
    name                    = 'Read',
    x                       = [row[0] for row in g_dir],
    y                       = [row[1] for row in g_dir],
    base                    = [0 for row in g_dir],
    offsetgroup             = 0,
    width                   = 0.8,
    marker_color            = C_GREEN), 
    go.Bar(
    name                    = 'Unread',
    x                       = [row[0] for row in g_dir],
    y                       = [row[2] for row in g_dir],
    base                    = [row[1] for row in g_dir],
    offsetgroup             = 0,
    width                   = 0.8,
    marker_color            = C_RED )])
fig.update_layout(barmode='stack')
fig.write_html('01.html', auto_open=False, full_html=False, default_width="100%", default_height="100%")

#-- Top senders
print("Finding top senders...")
topsenders                  = f_topsenders(inbox)
names                       = [item[0] for item in topsenders]
values                      = [item[1] for item in topsenders]
fig                         = go.Figure(data=[go.Bar(x=names, y=values, marker_color=C_BLUE)])
fig.write_html("02.html")

#-- Longest thread
print("Finding the longest thread...")
#longest_threads = f_longestthread(inbox)

#-- Top recipients
print("Finding top recipients where I am a sender...")
topsenders                  = f_toprecipients(sent)
names                       = [item[0] for item in topsenders]
values                      = [item[1] for item in topsenders]
fig                         = go.Figure(data=[go.Bar(x=names, y=values, marker_color=C_BLUE)])
fig.write_html("03.html")

#-- Top CC
print("Finding top senders putting me in CC...")
topsenders                  = f_topccsenders(inbox)
names                       = [item[0] for item in topsenders]
values                      = [item[1] for item in topsenders]
fig                         = go.Figure(data=[go.Bar(x=names, y=values, marker_color=C_BLUE)])
fig.write_html("04.html")

#-- Top mailing list
print("Finding top senders putting me in a mailing lists ...")
topsenders                  = f_topgroupsenders(inbox)
names                       = [item[0] for item in topsenders]
values                      = [item[1] for item in topsenders]
fig                         = go.Figure(data=[go.Bar(x=names, y=values, marker_color=C_BLUE)])
fig.write_html("05.html")

#-- Top meeting inviters
print("Finding top meeting inviters...")
g_top                       = f_topmeetinginviters(calendar)
names                       = list(g_top.keys())
values                      = list(g_top.values())
fig                         = go.Figure(data=[go.Bar(x=names, y=values, marker_color=C_BLUE)])
fig.write_html("06.html")

#-- Answers to invitations
print("Evaluating how I answer to meeting invitations...")
g_top                       = f_meetinganswers(calendar)
labels                      = [item[0] for item in g_top]
values                      = [item[1] for item in g_top]
colors                      = [C_GREEN, C_RED]
fig                         = go.Figure(data=[go.Pie(labels=labels, values=values, marker=dict(colors=colors))])
fig.update_traces(textinfo='label+value')
fig.update_traces(insidetextorientation='radial')
fig.write_html('07.html', auto_open=False, full_html=False, default_width="100%", default_height="100%")

#-- Time spent in meeting
print("Computing time spent in meetings...")
meetinghours, mymeetinghours = f_meetingtime(calendar)
meetingdays = int(meetinghours / 8)
mymeetingdays = int(mymeetinghours / 8)

#-- Conflicting meetings
print("Counting the number of conflicting meetings...")
nbconflict = f_conflictingmeetings(calendar)


longest_threads = [19, "Confidential: Project MKUltra"]
#-- Generating HTML
print("Generating HTML report...")
htmlstart = '''<hmtl><head><title>OutlookStats</title>
<style>a:link,
a:visited {
  color: inherit;
  text-decoration: none;
}body,html{margin:0;padding:0;width:100%;height:100%}.wrapper{width:100%;height:100%}footer,section{width:100%}.row{margin:auto;width:100%;max-width:60em}.col-2{display:inline-block;vertical-align:top}.col-2{width:100%}.row h1{margin:0;font-weight:300;font-size:40px}.sec-about{text-align:center}.row-grey{background:#111;color:#ddd}.sec-about .row-grey .row{padding:3em 0}.sec-news{text-align:center}.sec-news .row>h1{padding-bottom:.5em}.sec-news .col-2{padding:1em}.post{position:relative;height:700px;cursor:pointer;@extend %transition;box-shadow:0 0 10px rgba(0,0,0,.75)}.post h1,.post p,.post span{position:absolute}.post span{padding:.25em .5em;font-weight:700;color:#fff;opacity:.85}.post h1{bottom:0;width:100%;font-size:1.15em;line-height:2em;text-align:center;color:#fff;background:rgba(0,0,0,.75)}footer{background:#111}footer p{font-size:.85em;text-align:center;color:#aaa}</style>
<body><div class="wrapper"><section class="sec-about"><div class="row-grey"><div class="row"><article class="col-2">'''
htmlstart += '<h1>OutlookStats for ' + me + '</h1>Over the last ' + str(C_NBDAYS) + ' days</article></div></div></section><section class="sec-news">'
f = open(C_HTMLFILE, "w")
f.write(htmlstart)
f_htmlsection(f, C_REPORTS['readunread'])
f_htmlsection(f, C_REPORTS['topsenders'])
f_htmlsection(f, C_REPORTS['topccsenders'])
f_htmlsection(f, C_REPORTS['topgroupsenders'])
f_htmlsection(f, C_REPORTS['toprecipients'])
f.write(C_INFOSTART                 + 
    ''' <br><br><br><br><br><br><br><br>
    <h2 style="text-align: left; margin-left: 20px;">
    The longest thread is '''       + 
    str(longest_threads[0])         + 
    ''' emails long.<br><br>             
    The subject is <span style="color:red;margin-top: 
    0px;padding-top: 0px;">'''      + 
    longest_threads[1]              + 
    '''</span></h2><h1>'''          +
    C_REPORTS['longestthread'][1]   + 
    '''</h1>'''                     + 
    C_INFOEND)
f_htmlsection(f, C_REPORTS['topinviters'])
f_htmlsection(f, C_REPORTS['meetinganswers'])
f.write(C_INFOSTART                 + 
    '''
    <br><br><br><br><br><br>
    <h2 style="text-align: left; margin-left: 20px;">
    I have spent '''                +
    str(meetinghours)               + 
    '''</span> hours in meetings. <br><br>
    This the equivalent of '''      + 
    str(meetingdays)                +
    ''' days of work.
    <br><br><hr><br><br>
    I have setup '''                + 
    str(mymeetinghours)             + 
    ''' hours of meetings.<br><br>
    This the equivalent of '''      + 
    str(mymeetingdays)              + 
    ''' days of work. 
    <br><br><hr><br><br>
    And I had '''                   +
    str(nbconflict)                 +
    ''' meetings happening at the same time.
    </h2><h1>'''                    +
    C_REPORTS['meetings'][1]        +
    ''' </h1>'''                    +
    C_INFOEND)
f.write('</section><footer><div class="row"><p>&copy; 2023 <a href="https://www.linkedin.com/in/jld-ciso/" target="_blank">Jean-Luc Dupont</a></p></div></footer></div></body></html>')
f.close()
print("Launching report in browser...")
webbrowser.open('outlookstats.html', new=2)


""" ---------------------------------------------------------------------------
    End
    ----------------------------------------------------------------------- """
