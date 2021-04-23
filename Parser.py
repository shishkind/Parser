from tkinter import *
from tkinter.filedialog import *
import fileinput
import pyodbc
from pyodbc import Error
from lxml import etree

l = []
Full_results = []
gl_indicators = []
gl_ind_flags = []
source = ""
author_list = []
all_firstnames = []
all_lastnames = []
all_patr = []
xml=[]

def read():
    global l
    global source
    global xml
    root = Tk()
    root.withdraw()
    with open(askopenfilename(filetypes=[("Text files","*.xml *.bib")]), encoding='utf-8') as f:
        if str(f).find(".bib")!=-1:
            l = f.read().splitlines()
            for i in range(len(l)):  # Выявление источника экспорта путем поиска в файле адреса скопуса или категорий ВоС
                if l[i].find("https://www.scopus.com") != -1:
                    source = "Scopus"
                    break
                elif l[i].find("Web-of-Science") != -1:
                    source = "Web Of Science"
        if str(f).find(".xml")!=-1:
            xml = f.read()
            source = "eLibrary"
            #print(xml)
 
def parseXML():
    global xml
    root = etree.fromstring(xml)
    book_dict = []
    books = []
    indicators = []
    ind_dict = []
    for book in root.getchildren():
        for elem in book.getchildren():
          if len(elem.getchildren()) > 0 :
            for sub_elem in elem.getchildren():
              if len(sub_elem.getchildren()) > 0 :
                for sub_elem1 in sub_elem.getchildren():
                  if len(sub_elem1.getchildren()) > 0 :  
                    for sub_elem2 in sub_elem1.getchildren():
                      if len(sub_elem2.getchildren()) > 0 :  
                        for sub_elem3 in sub_elem2.getchildren():
                          if not sub_elem3.text:
                            text = "None"
                          else:
                            text = sub_elem3.text
                          #print(sub_elem.tag + " => " + text)
                          book_dict.append(text)
                          ind_dict.append(sub_elem3.tag)
                        continue
                      if not sub_elem2.text:
                        text = "None"
                      else:
                        text = sub_elem2.text
                      #print(sub_elem.tag + " => " + text)
                      book_dict.append(text)
                      ind_dict.append(sub_elem2.tag)
                    continue
                  if not sub_elem1.text:
                    text = "None"
                  else:
                    text = sub_elem1.text
                  #print(sub_elem.tag + " => " + text)
                  book_dict.append(text)
                  ind_dict.append(sub_elem1.tag)
                continue
              if not sub_elem.text:
                text = "None"
              else:
                text = sub_elem.text
              #print(sub_elem.tag + " => " + text)
              book_dict.append(text)
              ind_dict.append(sub_elem.tag)
            continue
          if not elem.text:
                text = "None"
          else:
              text = elem.text
          #print(elem.tag + " => " + text)
          book_dict.append(text)
          ind_dict.append(elem.tag)
        
        if book.tag == "item":
            books.append(book_dict)
            indicators.append(ind_dict)
            ind_dict = []
            book_dict = []

    for i in range(len(books)):
      for j in range(len(books[i])):
        print(indicators[i][j] + ": " + books[i][j])
      print(" ")
    #return books
 
 
def parse_WoS(n):
    global gl_indicators
    indicators = ['Author', 'Editor', 'Book-Group-Author', 'Title', 'Booktitle', 'Journal', 'Year', 'Volume', 'Number',
                  'Pages',
                  'Month', 'Note', 'Organization', 'Abstract', 'Publisher', 'Address', 'Type', 'Language',
                  'Affiliation', 'DOI', 'Early Access Date', 'ISSN', 'EISSN', 'ISBN', 'Keywords', 'Keywords-Plus',
                  'Research-Areas', 'Web-of-Science-Categories', 'Author-Email', 'ResearcherID-Numbers',
                  'ORCID-Numbers',
                  'Cited-References', 'Number-of-Cited-References', 'Times-Cited', 'Usage-Count-Last-180-days',
                  'Usage-Count-Since', 'Journal-ISO', 'Doc-Delivery-Number', 'Unique-ID', 'OA', 'DA']
    ind_flags = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
                 0, 0, 0, 0, 0, 0, 0]
    tmp_str = ""
    tmp = []
    result = ""
    i1 = 0
    results = []
    tmp_ind = 0
    flag = 0
    new_art = 0
    for i in range(n, len(l)):  # Построчная запись файла в tmp, пока не встретится пустая строка
        if l[i] != '':  # Если строка не пустая, то запись в тмп
            tmp.append(l[i])
        else:  # Если пустая, в качестве начала указывается адрес следующей статьи и цикл записи прерывается
            new_art = i + 2
            break
    while i1 < len(tmp):  # Построчная анализ блока
        if i1 == 0:  # Для первой записи отдельный обработчик, поскольку авторы в ВоС содержатся в одинарных скобках,
            # в отличии от остальных показателей
            if tmp[i1].find('}') == -1 and tmp[i1].find(
                    '{') != -1:  # Если в строке нет закрывющейся скобки, то к ней прибавляется следующая, после чего
                # следующая строка удаляется
                tmp[i1] = tmp[i1] + tmp[i1 + 1][2:]
                del tmp[i1 + 1]
                continue
            if tmp[i1].find('{') == -1 and tmp[i1].find(
                    '}') != -1:  # Если в строке  нет открывающейся скобки, то она прибавляется к предыдущей и удаляется
                tmp[i1 - 1] = tmp[i1 - 1] + tmp[i1]
                del tmp[i1]
                continue
            if tmp[i1].find('}') != -1 and tmp[i1].find(
                    '{') != -1:  # Если в строке есть обе скобки, то строка обрабатывается
                tmp_str = tmp[i1]
                if tmp_str.find('Author =') != -1:  # Проверка, что первая строка действительно содержит авторов
                    for c in range(len(tmp_str)):  # Поиск открытой и закрытой скобок
                        open = tmp_str.find('{')
                        close = tmp_str.find('}')
                        break
                    for c in range(open + 1, close):  # Посимвольное добавление текста в скобках в переменную
                        result = result + tmp_str[c]
                    results.append(result)  # Добавление перемнной в общий список
                    # print(result)
                    result = ''  # Обнуление переменной
                    ind_flags[0] = 1  # Установка флага
                i1 += 1  # Следующий шаг
                # print("")
        else:  # Для всех остальных показателей, которые содержатся в двойных скобках
            if tmp[i1].find('}}') == -1 and tmp[i1].find(
                    '{{') == -1:  # Если в строке нет скобок, то переход к следующему шагу
                i1 += 1
                continue
            if tmp[i1].find('}}') == -1 and tmp[i1].find('{{') != -1:  # Аналогично авторам
                tmp[i1] = tmp[i1] + tmp[i1 + 1][2:]
                del tmp[i1 + 1]
                continue
            if tmp[i1].find('{{') == -1 and tmp[i1].find('}}') != -1:  # Аналогично авторам
                tmp[i1 - 1] = tmp[i1 - 1] + tmp[i1]
                del tmp[i1]
                continue
            if tmp[i1].find('}}') != -1 and tmp[i1].find('{{') != -1:  # Аналогично авторам
                tmp_str = tmp[i1]
                for ind in range(tmp_ind, len(indicators)):  # Проход по списку индикаторов
                    if ind_flags[ind] == 1:  # Если флаг показателя уже установлен, то переход к следующему
                        continue
                    if tmp_str.startswith(indicators[ind]) != 0:  # Если строка начинается с элемента списка показателей
                        for c in range(len(tmp_str)):  # Поиск скобок
                            open = tmp_str.find('{') + 1
                            close = tmp_str.find('}}')
                            break
                        for c in range(open + 1, close):  # Посимвольная запись
                            result = result + tmp_str[c]
                        results.append(result)  # Добавление переменной в список
                        result = ''  # Обнуление
                        ind_flags[ind] = 1  # Установка флага
                        tmp_ind = ind  # переменная для установки диапазона цикла
                        break
                i1 += 1
            else:
                break
    gl_ind_flags.append(ind_flags)  # Добавление флагов статьи в общий список флагов
    Full_results.append(results)  # Добавление списка результатов в общий список
    if new_art == 0:
        gl_indicators = indicators
        return
    parse_WoS(new_art)


def parse_Scopus(n):  # Аналогично ВоС
    global gl_indicators
    indicators = ['author', 'title', 'journal', 'year', 'volume', 'number', 'pages', 'doi', 'art_number', 'note', 'url',
                  'affiliation',
                  'correspondence_address1', 'editor', 'publisher', 'issn', 'isbn', 'language', 'abbrev_source_title',
                  'document_type']
    ind_flags = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    new_art = 0
    tmp = []
    tmp_str = ''
    result = ''
    results = []
    tmp_ind = 0
    for i in range(n, len(l)):
        if l[i] != '':
            tmp.append(l[i])
        else:
            new_art = i + 1
            break
    for i in range(len(tmp)):
        tmp_str = tmp[i]
        for ind in range(tmp_ind, len(indicators)):
            if ind_flags[ind] == 1:
                continue
            if tmp_str.startswith(indicators[ind]) != 0:
                for c in range(len(tmp_str)):
                    open = tmp_str.find('{')
                    close = tmp_str.find('}')
                    break
                for c in range(open + 1, close):
                    result = result + tmp_str[c]
                results.append(result)
                result = ''
                ind_flags[ind] = 1
                tmp_ind = ind

    gl_ind_flags.append(ind_flags)
    Full_results.append(results)
    if new_art == 0:
        gl_indicators = indicators
        return

    parse_Scopus(new_art)


def assignment():
    global gl_indicators
    global Full_results
    global gl_ind_flags
    for i in range(len(gl_ind_flags)):
        for j in range(1, len(gl_ind_flags[i])):
            # Если флаг равен нулю, то на его индекс в результаты вставляется пробел
            if gl_ind_flags[i][j] == 0:
                Full_results[i].insert(j, " ")


def split():
    global author_list
    pg_st = []
    pg_end = []
    count_pages = []
    global gl_indicators
    result = ''
    tmp_flag = 0
    lastname = []
    tmp_lastname = ""
    global all_lastnames
    tmp_firstname = ""
    firstname = []
    global all_firstnames
    patr = []
    global all_patr
    #if source == "eLibrary":
        #for i in range()
    for k in range(len(Full_results)):
        tmp_authors = []
        if gl_ind_flags[k][0] == 1 and gl_indicators[0].upper() == "AUTHOR":  # Если есть автор
            tmp = Full_results[k][0]  # Cтрока с авторами записывается в переменную
            for i in range(len(tmp)):
                if tmp.find(" and ") != -1:  # Если в строке есть and
                    tmp_and = tmp.find(" and ")  # Запоминание индекса
                    tmp_authors.append(tmp[:tmp_and])  # добавление в список всего до индекса and
                    tmp = tmp[tmp_and + 5:]  # Присваивание переменной оставшуюся часть строки
                    result = ''
                else:
                    tmp_authors.append(tmp)  # Если в строке нет and, то она просто записывается в переменную
                    break
        author_list.append(tmp_authors)  # Добавление переменной в список

        for i in range(len(Full_results[k])):
            if gl_indicators[i].upper() == 'PAGES':  # Если в статье есть страницы
                if Full_results[k][i].find("-") != -1:  # Если есть разделитель
                    tmp_st = Full_results[k][i][:Full_results[k][i].find("-")]  # Запоминание части до разделителя
                    pg_st.append(tmp_st)
                    tmp_end = Full_results[k][i][
                              Full_results[k][i].find("-") + 1:]  # Запоминание части после разделителя
                    pg_end.append(tmp_end)
                    count_pages.append(int(tmp_end) - int(tmp_st) + 1)  # Количество страниц
                else:  # Если страница одна, то она записывается в качестве начальной и конечной
                    pg_st.append(Full_results[k][i])
                    pg_end.append(Full_results[k][i])
                    count_pages.append(1)
                tmp_flag = 1  # Флаг на то, что хотя бы в одной из статей есть страницы
            if gl_indicators[i].upper() == 'NOTE' and Full_results[k][i].find("cited By") != -1:  # есть в поле note есть количество цитирований, то они переформатируются просто в чисто
                Full_results[k][i] = Full_results[k][i][Full_results[k][i].find("cited By") + 9:]
            if gl_indicators[i].upper() == "ISSN" and Full_results[k][i].find("-") != -1:
                Full_results[k][i] = Full_results[k][i][:Full_results[k][i].find("-")] + Full_results[k][i][Full_results[k][i].find("-")+1:]

        for i in range(len(author_list[k])):  # Разделение фио автора
            # print(author_list[k][i])
            for j in range(len(author_list[k][i])):  # Посимвольое перебирание символов
                if author_list[k][i][j] != ',':  # Пока не встретится запятая
                    tmp_lastname = tmp_lastname + author_list[k][i][j]  # Посимвольная запись в переменную
                else:
                    author_list[k][i] = author_list[k][i][j + 2:]  # Если встретилась запятая, то переназначение переменной на строку после запятой и прерывание цикла
                    break
            for j in range(len(author_list[k][i])):
                if author_list[k][i][j] != ' ':  # Аналогично для имени, только вместо запятой ищется пробел
                    tmp_firstname = tmp_firstname + author_list[k][i][j]
                else:
                    author_list[k][i] = author_list[k][i][j + 1:]  # Аналогичное переназначение
                    break


            lastname.append(tmp_lastname)  # Добавление фамилии в общий список фамилий для одной статьи
            firstname.append(tmp_firstname)  # Аналогично, но для имен
            patr.append(author_list[k][i])  # Аналогично для отчества, но в качестве отчества берется остаток строки
            author_list[k][i] = []
            tmp_lastname = ""
            tmp_firstname = ""
        all_lastnames.append(lastname)  # Добавление фамилий авторов статьи в общий список фамилий для всего файла
        all_firstnames.append(firstname)  # Аналогично имена
        all_patr.append(patr)  # Фамилии туда же
        lastname = []
        firstname = []
        patr = []

    if tmp_flag == 1:  # Добавление новых показателей, флагов и инфы о страницах в общие списки (Проверка по флагу и прошлого блока)
        for i in range(len(gl_indicators)):  # Построчный перебор индикаторов
            if gl_indicators[i].upper() == 'PAGES':  # Если нашли в строке pages
                gl_indicators.insert(i, "Count of pages")  # Вставка на этот индекс новых показателей
                gl_indicators.insert(i, "End page")
                gl_indicators.insert(i, "Start page")
                del gl_indicators[i + 3]  # Удаление показателя pages
                for j in range(len(Full_results)):  # Аналогичное действие для обработанных данных
                    Full_results[j].insert(i, str(count_pages[j]))
                    Full_results[j].insert(i, pg_end[j])
                    Full_results[j].insert(i, pg_st[j])
                    del Full_results[j][i + 3]
                    gl_ind_flags[j].insert(i, 1)  # Для флагов тоже
                    gl_ind_flags[j].insert(i, 1)
                    gl_ind_flags[j].insert(i, 1)
                    del gl_ind_flags[j][i + 3]


def output(): # Тут просто вывод общего списка  показателей (уже допиленного), для авторов отдельный вывод
    global source
    global gl_indicators
    global Full_results
    global gl_ind_flags
    global all_patr
    global all_lastnames
    global all_firstnames
    #if source == "eLibrary":
    for i in range(len(Full_results)):
        if gl_ind_flags[i][0] == 1 or gl_indicators[0].upper == "AUTHOR":
            for j in range(len(author_list[i])):
                print("Author " + str(j + 1) + ":  " + all_lastnames[i][j] + ", " + all_firstnames[i][j] + " " +
                      all_patr[i][j])
        for j in range(1, len(gl_ind_flags[i])):
            print(str(j) + " " + gl_indicators[j] + ":  " + Full_results[i][j])
        print("")
    print("Exported from " + source)




def sql_output():
    global all_firstnames
    global all_patr
    global all_lastnames
    server = '192.168.20.170'
    database = 'BasePPS'
    username = 'shishkin.d'
    password = '295#SKz'
    try:
        cnxn = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
        cursor = cnxn.cursor()
        print("Successful connection to db")
    except Error as e:
        print("Connection Error!")
        return
    for k in range(len(Full_results)):
        UIDs = []
        issn_flag = 0
        cursor.execute("""SELECT UID, FirstName, Patronymic, LastName FROM dbo.Author""")
        Author_data = cursor.fetchall()
        cursor.execute("""SELECT structural_part_id_part FROM dbo.Author_publ""")
        struct_tmp = cursor.fetchall()
        print(struct_tmp)
        if len(struct_tmp) > 0:
            struct_id = max(struct_tmp)
        else:
            struct_id = 0
        for i in range(len(all_lastnames[k])):
            UIDs.append("")
            for j in range(len(Author_data)):
                if all_lastnames[k][i] == Author_data[j][3]:
                    UIDs[i] = Author_data[j][0]
                #else:
                    #UIDs[i] = data[len(data[j])-1][0]
        print(UIDs)
        for i in range(len(UIDs)):
            if UIDs[i] == "":
                UIDs[i] = len(Author_data) + i + 1
        for i in range(len(Author_data)):
            print(Author_data[i])
        print(UIDs)
        for i in range(len(all_lastnames[k])):
            if int(UIDs[i]) > len(Author_data):
                cursor.execute("""INSERT INTO dbo.Author(UID, FirstName, Patronymic, LastName, Office_depart_kod_office) VALUES(?,?,?,?,?)""", UIDs[i], all_firstnames[k][i], all_patr[k][i], all_lastnames[k][i], '037')
                cnxn.commit()
        cursor.execute("""SELECT ISSN FROM dbo.Series""")
        issn_check = cursor.fetchall()
        print(issn_check)
        for i in range(len(issn_check)):
            if Full_results[k][23] == issn_check[i][0]:
                print("+++++")
                issn_flag = 1
        if issn_flag == 0:
            cursor.execute("""INSERT INTO dbo.Series (ISSN, ISSN_O, Volume, Dare_Edit, Date_ExtractJ, RELEASE, User_id) VALUES(?,?,?,?,?,?,?)""", Full_results[k][23], Full_results[k][23], Full_results[k][7], "2020-01-01", "2020-01-01", 1, 1)
        cnxn.commit()
        cursor.execute("""INSERT INTO dbo.edition (id_Edt, Name_Edt, YEAR, DOI_ED, publishing_office_id_PBO, Spr_formatInfo_Type_ED, SERIES_ISSN, id_language, id_edition) VALUES(?,?,?,?,?,?,?,?,?)""", k, Full_results[k][5], Full_results[k][6], Full_results[k][21], 1, 1, Full_results[k][23], 1, 1)
        cnxn.commit()
        cursor.execute("""INSERT INTO dbo.structural_part(id_part, Name_part, PageBg, PageEnd, PageCount, Date_Ins, User_Ins, Copy_id, Date_EdA, User_EdA, Date_ExtractA, edition_id_Edt, Spr_structure_id_TypePart) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", k, "temp title",  Full_results[k][9], Full_results[k][10], Full_results[k][11], "01-01-2020", k, k, "01-01-2020", k, "01-01-2020", k, 1)
        cnxn.commit()
        for i in range(len(all_lastnames[k])):
            cursor.execute("""INSERT INTO dbo.Author_publ(Authors_PartA, structural_part_id_part, Author_UID) VALUES(?,?,?)""", all_lastnames[k][i], k, UIDs[i])
            cnxn.commit()
        print(UIDs)
        print(struct_id)
    #except Error as e:
     #   print("Insert Error!")


def main():
    read()
    global source
    global l
    global Full_results
    global gl_ind_flags
    Full_results = []
    gl_ind_flags = []
    if source == "Scopus":
        parse_Scopus(3)
    elif source == "Web Of Science":
        parse_WoS(2)
    elif source == "eLibrary":
        parseXML()
    else:
        print("File Error!")
        return
    assignment()
    split()
    output()
    #if source == "Web Of Science":
        #sql_output()
    #else:
        #print("sql для скопуса в разработке")


if __name__ == '__main__':
    main()

