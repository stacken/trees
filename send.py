#! /usr/bin/env python
# -*- coding: utf-8 -*-
import smtplib
import codecs
from email.mime.text import MIMEText
from email.header import Header
import fileinput
import datetime

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

    if reply_to:
        msg['Reply-to'] = reply_to

    if smtp:
        print "[{}] Sending to {}".format(pos, msg['To'])
        smtp.sendmail(fromaddr, toaddrs, msg.as_string())
    else:
        print "[{}] Not sending to {}".format(pos, msg['To'])

from optparse import OptionParser
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
    addressre = re.compile('(.*)(<[A-Za-z0-9._]+@[a-z.]+>)')
    for line in fileinput.input(options.addressfile,
                                openhook=fileinput.hook_encoded('utf-8')):
        line = line.rstrip()
        m = addressre.match(line);
        if line == '':
            pass
        elif m:
            msg = mkmsg(msgfile[0], subject=options.subject,
                        fromname=u'Datorföreningen Stacken via {0}'.format(options.from_name),
                        fromaddr='<{0}>'.format(options.from_email),
                        toname = m.group(1),
                        toaddrs = [m.group(2)],
                        reply_to = options.reply_to,
                        smtp = server)
        else:
            print "Bad line: " + line

elif options.finger:
    print 'Reading people from finger.txt'
    n = 0;
    import re
    kontonu = re.compile('^[a-z_0-9]+$')
    kontosen = re.compile('^\([a-z_0-9/,]+\)$')
    for line in fileinput.input('out/finger.txt',
                                openhook=fileinput.hook_encoded('iso-8859-1')):
        fields = line.rstrip().split(';')
        ( efternamn, fornamn, sortering, titel, c_o, 
          avdelning, organisation, gatuadress, postadress, land, distribution,
          hemtelefon, arbtelefon,
          ppn, anvandarnamn, mailadress, betalt, intradesdatum, uttradesdatum, 
          status, kortnr, xx) = fields + (22-len(fields))*[None]

        if (status==None) or (efternamn==u'Efternamn' and fornamn==u'Förnamn' and kortnr==u'Kortnr'):
            continue

        flags = re.split(',\s*', status)

        if betalt: betalt = int(betalt)
        else:      betalt = 0

        if ((betalt < datetime.datetime.now().year - 3)
            and not ('Ny' in flags)
            and not ('Hedersmedlem' in flags)):
            continue

        addrs = []
        if anvandarnamn:
            if kontonu.match(anvandarnamn):
                addrs = addrs + ['<' + anvandarnamn + '@stacken.kth.se>']
            elif not kontosen.match(anvandarnamn):
                print u'!!! Strange account name: ' + anvandarnamn

        if mailadress: addrs = addrs + ['<' + mailadress + '>']

        if (uttradesdatum
            or ('Slutat' in flags)
            or ('Utesluten' in flags)
            or ('Ej medlem' in flags)) :
            print u'!!! %s %s lämnade stacken %s (%s)' % (fornamn, efternamn, uttradesdatum, status)
            continue

        if len(addrs) < 1:
            print u"!!! No address found for %s %s" % (fornamn, efternamn)
            continue

        if (n < int(options.resume_from)):
            continue

        msg = mkmsg(msgfile[0], subject=options.subject,
                    fromname=u'Datorföreningen Stacken via {0}'.format(options.from_name),
                    fromaddr='<{0}>'.format(options.from_email),
                    toname = u'%s %s' % (fornamn, efternamn),
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
