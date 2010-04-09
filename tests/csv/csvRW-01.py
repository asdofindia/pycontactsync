import csv

rowdetail=1

ifile  = open('data-set.csv', "rb")
reader = csv.reader(ifile, delimiter='\t', quotechar='"', quoting=csv.QUOTE_ALL)

ofile  = open('output.csv', "wb")
writer = csv.writer(ofile, delimiter='\t', quotechar='"', quoting=csv.QUOTE_ALL)

rownum = 0
for row in reader:
    # Save header row.
    if rownum == 0:
        header = row
        writer.writerow(header)
        print header
    else:
        writer.writerow(row)

        if rowdetail:
            print row
        else:
            colnum = 0
            for col in row:
                print '%-8s: %s' % (header[colnum], col)
                colnum += 1
   
    print 'Row number %d: ' % rownum
    rownum += 1

ifile.close()
ofile.close()
