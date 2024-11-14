import pandas as pd
import json
import logging
import os

log_file = 'logs/app.log'  #setting log file for all logs to be stored
ifp = 'data/excelfile.xlsx'  #input file path
ofp = 'data/file.json'  #output file path for the generated json file
courses = []  #list with parsed course data

#logging setup
logging.basicConfig(filename=log_file, level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

#json dump
def paste(courses, ofp):
    try:
        with open(ofp, 'w') as f:
            json.dump(courses, f, indent=4)
    except Exception as e:
        print(f"crror occurred: {e}")
    print(f"created json at {ofp}")


#to go through excel file via pandas library and collect information in the structure required
def parse(ifp, courses):
    logging.info(f"beginning to parse file at {ifp}")
    try:
        xls = pd.ExcelFile(ifp)
        for sheet in xls.sheet_names:
            course_data = {}
            instructors = []  
            df = pd.read_excel(ifp, sheet_name=sheet, header=[1])
            
            #to extract common info for whoole course
            course_code = df.iloc[1, 1] if pd.notna(df.iloc[1, 1]) else ""
            course_title = df.iloc[1, 2] if pd.notna(df.iloc[1, 2]) else ""
            credits = {
                "lecture": df.iloc[1, 3] if pd.notna(df.iloc[1, 3]) else 0,
                "practicals": df.iloc[1, 4] if pd.notna(df.iloc[1, 4]) else 0,
                "units": df.iloc[1, 5] if pd.notna(df.iloc[1, 5]) else 0
            }
            
            sections = []  
            mid = ''
            compre = ''
            
            #iterating through each row to select section specific data
            for index, row in df.iterrows():
                section_code = df.iloc[index, 6] 
                if pd.notna(section_code): 
                    if len(sections) > 0:
                        sections[-1]["instructors"] = instructors  
                    instructors = [] 
                    instructors.append(df.iloc[index,7])

                    #finding out section type
                    sectype = ''
                    if section_code[0] == 'L':
                        sectype = 'lecture'
                    elif section_code[0] == 'T':
                        sectype = 'tutorial'
                    elif section_code[0] == 'P':
                        sectype = 'practical'
                    
                    #arranging timing information from M W 2 format to days and slots format
                    timing = []
                    dh = df.iloc[index, 9]  
                    if pd.notna(dh):
                        for i in range(len(dh)):
                            if dh[i].isdigit():  
                                temp = dh[:i]
                                slot = [int(dh[i]), int(dh[i]) + 1]  
                                for a in temp:
                                    if a == "T" and temp[temp.index(a) + 1] == "h":
                                        timing.append({"day": "Th", "slots": slot})
                                    elif a.isalpha()==True: 
                                        timing.append({"day": a, "slots": slot})
                    
                    #make a new section with collected section data and append to list
                    section_data = {
                        "section_type": sectype,
                        "section_number": section_code,
                        "instructors": instructors,
                        "room": str(int(df.iloc[index, 8])),  
                        "timing": timing
                    }
                    sections.append(section_data)
                
                #get dates for midsems and compres
                if pd.notna(df.iloc[index, 10]) and not mid and not compre:
                    mid = str(df.iloc[index, 10]) 
                    compre = str(df.iloc[index, 11]) 
                
                #adding instructors to current sections
                if pd.isna(section_code) and len(sections) > 0:
                    instructor = df.iloc[index, 7]  
                    if pd.notna(instructor):
                        sections[-1]["instructors"].append(instructor)

            #compiling data for each course to add to final courses list
            course_data = {
                "course_code": course_code,
                "course_title": course_title,
                "credits": credits,
                "sections": sections,
                "midsem date & session": mid,
                "compre date & session": compre
            }
            courses.append(course_data)  
        
        logging.info(f"parsed {len(courses)} courses.")
        return courses

    except Exception as e:
        logging.error(f"error parsing file at {ifp}: {e}")
        raise
    
#main function
def main():
    logging.info("running script to parse timetable data")
    try:
        parsed_courses = parse(ifp, courses)  
        paste(courses, ofp) 
        logging.info("excel parsed, json generated")
        exit(0) 
    except Exception as e:
        logging.error(f"error running script: {e}")
        exit(1)

if __name__ == '__main__':
    main()
