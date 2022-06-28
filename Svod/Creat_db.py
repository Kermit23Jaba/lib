import sqlite3

with sqlite3.connect('C:/Users/user/Desktop/svod_e.db') as con:

    cur = con.cursor()

    # cur.execute("""CREATE TABLE IF NOT EXISTS svod_avk_elements (
    #     Line TEXT NOT NULL,
    #     Original_Number_AVK INTEGER NOT NULL,
    #     Name_AVK TEXT NOT NULL,
    #     Request TEXT NOT NULL,
    #     Number INTEGER NOT NULL,
    #     Material TEXT NOT NULL,
    #     Mark TEXT NOT NULL,
    #     Plavka TEXT NOT NULL,
    #     Quantity TEXT NOT NULL,
    #     Certificates TEXT NOT NULL,
    #     Number_of_pages INTEGER NOT NULL
    # )""")


cur.executemany("INSERT INTO svod_avk_elements VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", tab)

con.commit





###########################################
import sqlite3
 
con = sqlite3.connect('C:/Users/user/Desktop/svod_e.db')
 
def sql_insert(con, entities):
 
    cursorObj = con.cursor()
    
    cursorObj.execute('INSERT INTO svod_avk_elements VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', entities)
    
    con.commit()
 
 
sql_insert(con, entities)
