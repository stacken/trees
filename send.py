#! /usr/bin/env python
# -*- coding: utf-8 -*-
import smtplib
import codecs
from email.mime.text import MIMEText
from email.header import Header
import fileinput
import datetime
import json
import sys

if not sys.stdout.encoding:
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout)

def mkmsg(filename, subject, fromname, fromaddr, toname, toaddrs,
          smtp, reply_to = None, outcs = 'iso-8859-1', pos=0):

    f = codecs.open(filename, mode='r', encoding='utf-8')
    msg = MIMEText(f.read().encode(outcs), 'plain', outcs)
    f.close()

    msg['Subject'] = Header(subject.decode('utf-8').encode(outcs), outcs)
    msg['From'] = (Header(fromname.encode(outcs), outcs).__str__()
                   + fromaddr)
    msg['To'] = (Header(toname.encode(outcs), outcs).__str__()
                 + ' ' + ', '.join(toaddrs))
    msg['X-Stacken'] = "There's a coffee bug living in the club room."

    if reply_to:
        msg['Reply-to'] = reply_to

    if smtp:
        print "[{}] Sending to {}".format(pos, msg['To'])
        smtp.sendmail(fromaddr, toaddrs, msg.as_string())
    else:
        print "[{}] Not sending to {}".format(pos, msg['To'])

from optparse import OptionParser
def main():
    opt = OptionParser(usage='usage: %prog [options] msgfile')
    opt.add_option('--doit', action='store_true', dest='doit', default=False,
                help='Actually do send messages')
    opt.add_option('--addressfile', dest='addressfile', metavar='FILE',
                help='File containing addresses to send to')
    opt.add_option('--finger', action='store_true', dest='finger', default=False,
                help='Find suitable recipients in finger.txt')
    opt.add_option('--subject', dest='subject',
                help='Subject text for the messages')
    opt.add_option('--reply-to', dest='reply_to', metavar='ADDR',
                help='Address for a Reply-to header')
    opt.add_option('--from-email', dest='from_email', metavar='EMAIL', help='Your E-mail adress')
    opt.add_option('--from-name', dest='from_name', metavar='NAME', help='Your Full Name')
    opt.add_option('--resume-from', dest='resume_from', metavar='NUMBER',
                help='Resume from this position (only for finger)', default=0)
    (options, msgfile) = opt.parse_args()

    if not options.subject:
        raise Exception, 'A subject is required'

    if not options.from_email:
        raise Exception, 'From E-mail is required'

    if not options.from_name:
        raise Exception, 'From Name is required'

    if not msgfile or len(msgfile) != 1:
        raise Exception, 'Exactly one message file is required'

    server = None
    if options.doit:
        print 'Initializing smtp'
        server = smtplib.SMTP('smtp.stacken.kth.se')
        server.set_debuglevel(0)

    if options.addressfile:
        print 'Reading addresses from ' + options.addressfile
        import re
        addressre = re.compile('(.*)(<[A-Za-z0-9._-]+@[A-Za-z0-9.-]+>)')
        for line in fileinput.input(options.addressfile,
                                    openhook=fileinput.hook_encoded('utf-8')):
            line = line.rstrip()
            m = addressre.match(line);
            if line == '':
                pass
            elif m:
                msg = mkmsg(msgfile[0], subject=options.subject,
                            fromname=u'Datorföreningen Stacken via {0}'.format(options.from_name.decode('utf-8')),
                            fromaddr='<{0}>'.format(options.from_email),
                            toname = m.group(1),
                            toaddrs = [m.group(2)],
                            reply_to = options.reply_to,
                            smtp = server)
            else:
                print "Bad line: " + line

    elif options.finger:
        print 'Reading people from finger.txt'
        n = 0
        import re
        kontonu = re.compile('^[a-z_0-9]+$')
        kontosen = re.compile('^\([a-z_0-9/,]+\)$')

        with open('/afs/stacken.kth.se/home/stacken/Private/finger_txt/finger.json') as json_data:
            for user in json.load(json_data):

                # Gör alla nycklar lowercase
                user = {k.lower():v for k,v in user.items()}

                # Hoppa över om vi inte har fått en betalning de senaste åren,
                # men inte om Ny eller Hedersmedlem.
                thisyear = datetime.datetime.now().year
                if ((user.get('betalt', 0) < thisyear - 3)
                    and (user.get('THS-studerande', 0) < thisyear - 3)
                    and not (user.get('ny', False))
                    and not (user.get('hedersmedlem', None))):
                    continue

                # Bygg en lista av epost-adresser att skicka till.
                addrs = []
                if user.get(u'användarnamn', None):
                    if kontonu.match(user[u'användarnamn']):
                        addrs.append('<' + user[u'användarnamn'] + '@stacken.kth.se>')
                    elif not kontosen.match(user[u'användarnamn']):
                        print u'!!! Strange account name: ' + user[u'användarnamn']

                if user.get('mailadress', None):
                    if not '<' + user['mailadress'] + '>' in addrs:
                        addrs.append('<' + user['mailadress'] + '>')

                # Hoppa över om vi inte hittade någon epost-adress.
                if len(addrs) < 1:
                    print u"!!! No address found for {} {}".format(
                            user.get(u'förnamn', 'No first name'),
                            user.get(u'efternamn', 'No last name')
                        )
                    continue

                # Hämta flaggor från status-fältet, utträdesdatum, slutat,
                # utesluten, hedersmedlem, och ny har egna fällt.
                flags = re.split(',\s*', user.get('status', ""))

                # Hoppa över personer som har slutat
                if (user.get(u'utträdesdatum', None)
                    or (user.get('slutat', False))
                    or (user.get('utesluten', False))
                    or ('Ej medlem' in flags)):
                    print u'!!! {} {} lämnade stacken {}'.format(
                            user.get(u'förnamn', 'No first name'),
                            user.get(u'efternamn', 'No last name'),
                            user.get(u'utträdesdatum', 'Inget utträdesdatum')
                        )
                    continue

                # Hoppa över personer som har bett om att inte få mailutskick
                if not user.get(u'epost-utskick', True):
                    print u'!! {} has requested that they do not receive e-mails'.format(user.get(u'användarnamn'))
                    continue

                # Säkerställ att det finns ett för och efternamn
                if not user.get(u'förnamn', None) and not user.get(u'efternamn', None):
                    print u'!! Missing name for user {}'.format(user.get(u'användarnamn', 'No username'))
                    continue

                # --resume-from=N
                if (n < int(options.resume_from)):
                    n = n + 1
                    continue

                msg = mkmsg(msgfile[0], subject=options.subject,
                            fromname=u'Datorföreningen Stacken via {0}'.format(options.from_name.decode('utf-8')),
                            fromaddr='<{0}>'.format(options.from_email),
                            toname = u'%s %s' % (user.get(u'förnamn', ''), user.get(u'efternamn', '')),
                            toaddrs = addrs,
                            reply_to = options.reply_to,
                            smtp = server,
                            pos=n)

                n = n + 1

        print 'There was %s people to send to.' % n

    else:
        print 'No addresses.  Sending to myself for debugging.'
        msg = mkmsg(msgfile[0], subject=options.subject,
                    fromname=u'Datorföreningen Stacken via {0}'.format(options.from_name),
                    fromaddr='<{0}>'.format(options.from_email),
                    toname = u'{0}'.format(options.from_name),
                    toaddrs  = ['<{0}>'.format(options.from_email)],
                    reply_to = options.reply_to,
                    smtp = server)

    if server:
        server.quit()

if __name__ == "__main__":
    main()
