from docx import Document
import csv

def raw_text_gen(file):
    '''
    :param file: Enter docx file to be extracted from
    :return: Returns raw data from docx file
    '''
    document = Document(file)
    raw_text = []
    for p in document.paragraphs:
        twips = p.paragraph_format.space_before
        if (twips is not None) and (13 <= twips * 0.000079):
            raw_text = raw_text + ['x', 'x', 'x', 'x', 'x']
        raw_text.append(p.text)
    return raw_text

def special_super_sort(raw_text, cent_word, side_word, lead, follow, sides):
    '''
    :param raw_text: Enter an array which contains all the raw text
    :param cent_word: Enter an array which contains all the words found within the centers
    :param side_word: Enter an array which contains all the words found within the sides
    :param side_word: Enter an array which contains all the words to be removed from sides
    :param lead: Enter text you do not want leading side words
    :param follow: Enter text you do not want following side words
    :param sides: Enter true or false [T/F] based on how you want sides to be displayed
    :return: Returns an array of the centers (not center words) and sides/side words
    '''
    raw_side_ind = []
    raw_side_word = []
    for a in range(len(raw_text)):
        for b in side_word:
            if (b in raw_text[a]) and ((lead+b+follow) not in raw_text[a]):
                raw_side_ind.append(a)
                raw_side_word.append(b)
                break
    cent = []
    side = []
    for c in range(len(raw_text)):
        for d in cent_word:
            if (d in raw_text[c]) and (len(raw_text[c]) < 50):
                cent.append(raw_text[c].strip())
                f = [abs(c - e) for e in raw_side_ind]
                g = f.index(min((g for g in f if g > 0)))
                if sides == 'F':
                    side.append(raw_side_word[g])
                else:
                    sep = raw_side_word[g]
                    side.append(raw_text[raw_side_ind[g]].split(sep, 1)[0] + sep)
                break
    return cent, side

def corrector(array, dict):
    '''
    :param array: Enter array to be corrected
    :param dict: Enter dictionary relating incorrect words to correct words
    :return: Returns corrected array
    '''
    new_array = []
    for g in array:
        try:
            if g in dict.keys():
                new_array.append(dict[g])
            else:
                for key in dict.keys():
                    new_array.append(g.replace(key, dict[key]))
                    break

        except:
            new_array.append(g)
    return new_array

def csv_matrix(file):
    '''
    :param file: Enter CSV file name
    :return: Return matrix of CSV data
    '''
    csv_id = (open(file, 'r'))
    csv_matrix = []
    csv_info = csv.reader(csv_id)
    for row in csv_info:
        csv_matrix.append(row)
    csv_id.close()
    return csv_matrix

raw_text = raw_text_gen('Medications_test.docx')

med_word = ['TAB', 'INJ', 'BAG', 'SOLN', 'CAP', 'DUP', 'LIQUID', 'Soln']

path_word = ['ORALLY', 'SUBCUT', 'IVPB', 'IVPUSH', 'IV', 'TOPICAL']

freq_word = ['BID', 'Q12H', 'Q6HPRN', 'DAILY', 'Q24H', 'Q4H', 'NOW', 'Q6', 'Q12',
             '012H', '012', '6HPRN', '024H', '06', '24H']

freq_word_error = {'012H':'Q12H', '012':'Q12', '6HPRN':'Q6HPRN', '024H':'Q24H',
                   '06':'Q6', '24H':'Q4H'}

day_word = ['4-09-20', '4-10-20', '4-11-20', '4-12-20', '4-13-20', '4-14-20',
            '4-15-20', '4-16-20', '4-17-20']

unit_word = ['MG', 'MEQ', 'ML', 'MEO']

unit_word_error = {'MEO':'MEQ'}

matrix = csv_matrix('Jeenu_Medication.csv')
CSV_med_word = [matrix[k][0] for k in range(len(matrix))]

med, path = special_super_sort(raw_text, med_word, path_word, 'sdflk', 'sdfhs', 'F')
a, freq = special_super_sort(raw_text, med_word, freq_word, 'sdfjk', 'sdfjsl', 'F')
b, day = special_super_sort(raw_text, med_word, day_word, 'gh ', '', 'F')
c, unit = special_super_sort(raw_text, med_word, unit_word, '', '/', 'T')

freq = corrector(freq, freq_word_error)
unit = corrector(unit, unit_word_error)

med_day = [med[z]+' x '+day[z] for z in range(len(med))]
no_dup = list(dict.fromkeys(med_day))
no_dup_ind = [med_day.index(y) for y in no_dup]

final_csv_id = open('Final_CSV_Data.csv', 'w')
writer = csv.writer(final_csv_id)

writer.writerow(['Day', 'Medication', 'Dose', 'Alternate Name', 'Function', 'Pathway', 'Frequency'])
for x in no_dup_ind:
    alt = ''
    app = ''
    prt = True
    for j in range(len(CSV_med_word)):
        if CSV_med_word[j] in med[x]:
            alt = matrix[j][1]
            app = matrix[j][2]
            if matrix[j][3] == 'NO':
                prt = False
            break
    if prt == True:
        writer.writerow([day[x], med[x], unit[x], alt, app, path[x], freq[x]])
final_csv_id.close()

