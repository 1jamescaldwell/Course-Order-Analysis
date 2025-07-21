'''
James Caldwell
UVA IRA
5/7/24

This script analyzes the impact of course-taking order on student GPA at the University of Virginia. It was created in response to a question about whether students perform differently depending on the sequence in which they take certain key classes —e.g., is it better to take Calculus before Linear Algebra, or the other way around?

The code is a query-style script with plans to integrate into UVA's UBI (university business intelligence). The user inputs a list of courses to analyze, and the script outputs a violin plot which visualizes how students have historically performed in those courses with respect to the order that a course is taken.

In the output plot, each course is followed by a number (1, 2, or 3 for example). The number corresponds to the order in the sequence in which the student took a course. If a student has two courses with the same number, that means they took those courses in the same semester.

To do:
deal with withdraws
Add notes about how 1 and 1 means simultaneous
Add error check if no students took all courses in the list
Add second plot to show without withdraws, and add withdraw count onto plot
'''

import pandas as pd
import numpy as np
import ast 
import os
import sys
import datetime
import openpyxl


## Plotting section for creating plots in VS code. Commented out for running with Qlik/automation.
# from matplotlib import pyplot as plt
# import seaborn as sns
# def create_violin_plot(plot_data):
#     fig, ax = plt.subplots(figsize=(10, 6))
#     sns.violinplot(
#         y='Course|Order w/ Count',
#         x='Avg Grade',
#         data=plot_data,
#         cut=0,
#         ax=ax
#     )
#     ax.set_ylabel("Course Order")
#     ax.set_xlabel("Avg Grade of Courses")
#     ax.set_title("Distributions of Grade by Class Order")
#     ax.set_xlim(0, 4)  # Optional: clip to specific grade range
#     return fig

def has_duplicate_course(course_list_str):
    '''
    Called by analyze_course_order, This function is used if include_repeats = False
    This function checks if there are duplicate courses in the course list string 
    It will cause a row that has entries like "BIOL 3000 1, BIOL 3000 1" to be dropped (i.e. a student re-took a course)
    '''
    try:
        course_list = ast.literal_eval(course_list_str)  # safely parse the string to a list
        courses_only = [c.rsplit(' ', 1)[0] for c in course_list]  # split off order number
        return len(set(courses_only)) < len(courses_only)  # True if duplicates exist
    except Exception:
        return True  # drop malformed entries

# Main function
def analyze_course_order(course_df, course_list,min_cutoff_course_number=50,include_repeats=False):
    '''This function calculates how many students have taken each of the classes under review 
    and analyzes how they performed depending on which order they took the classes.
    Inputs: entire UVA course grade dataset, course list (ex: [BIOL 3000, BIOL 3020]), and some parameters for plotting
    '''
    
    course_df['Course'] = course_df['Subject'] + ' ' + course_df['Catalog Number'].astype(str)
    course_df.drop(columns=['Subject', 'Catalog Number','Term Desc'], inplace=True)

    # Filter data to students who took these courses
    course_df.query('CourseAndTerm in @course_list', inplace=True)

    # Remove term desc from course_list. We already querried for the term desc, so we don't need it anymore.
    course_list = [c.split(' - ')[0] for c in course_list]

    # How many of each course was taken, regardless of if a student took all of them or not
    course_counts_initial = course_df['Course'].value_counts()

    course_df.sort_values(by=['Student System ID', 'Term'], inplace=True)
    # Add a column for the order of the course taken by each student
    course_df['Order'] = course_df.groupby('Student System ID')['Term'].rank(method='min').astype(int)

    course_df = course_df[course_df['Official Grade'] != 'W']  # Remove withdraws. Revisit withdraws at a later time

    # Replace grades with scale
    replace_map = {
        'A+':4.0,
        'A': 4.0,
        'A-': 3.7,
        'B+': 3.3,
        'B': 3.0,
        'B-': 2.7,
        'C+': 2.3,
        'C': 2.0,
        'C-': 1.7,
        'D+': 1.3,
        'D': 1.0,
        'D-': 0.7,
        'F': 0.0}
    course_df['Official Grade'] = course_df['Official Grade'].map(replace_map)

    # Combine 3 columns into 1
    course_df['Course|Order|Grade'] = course_df['Course'] + ' ' + course_df['Order'].astype(str) + ' ' + course_df['Official Grade'].astype(str)

    # Drop unnecessary columns
    keep_cols = ['Student System ID', 'Course|Order|Grade']
    course_df = course_df[keep_cols].copy()

    # Group courses by student, so each student has one row and Course|Order|Grade is a list of all courses taken by that student
    course_grouped = course_df.groupby('Student System ID')['Course|Order|Grade'].apply(list).reset_index()

    # Filter to students who took all courses in course_list
    query = course_grouped[
        course_grouped['Course|Order|Grade'].apply(
            lambda course_list_full: set(course_list).issubset(
                set(' '.join(item.split()[:2]) for item in course_list_full)
            )
        )
    ]
    # Create a new column with just the course and order without grade
    query = query.copy() # Make an explicit copy to avoid modifying the original / avoid copy warnings
    query.loc[:, 'Course|Order'] = query['Course|Order|Grade'].apply(
        lambda x: [' '.join(item.split()[:3]) for item in x]
    )

    # Create a new column with just the grades. It takes an average of the grades for each courses taken by the student from the list
        # If two courses are in the course search list, Avg Grade will be the mean of those two grades
    query.loc[:, 'Avg Grade'] = query['Course|Order|Grade'].apply(
        lambda x: sum([float(item.split()[3])/len(x) for item in x])  # Extract and add the grade values (the fourth element)
    )
    # Convert to string for grouping
    query['Course|Order'] = query['Course|Order'].astype(str)

    # Filter: Keep only rows where there are NO duplicates
        # It will cause a row that has entries like "BIOL 3000 1, BIOL 3000 1" to be dropped (i.e. a student re-took a course)
    if not include_repeats:
        query = query[~query['Course|Order'].apply(has_duplicate_course)].copy()

    min_cases = min_cutoff_course_number

    # Filter out courses with fewer than min_cases occurrences
    # Also calculate how many cases each instance has to display on the violin plot
    if query.shape[0] < 1: # if query was valid but just resulted in no students having taken that combination, return empty dataframes
        plot_data = pd.DataFrame()
        agg_result = pd.DataFrame()
    else:
        course_counts = query['Course|Order'].value_counts()
        print(course_counts)
        valid_courses = course_counts[course_counts >= min_cases].index
        plot_data = query[query['Course|Order'].isin(valid_courses)].copy()
        plot_data['Course|Order w/ Count'] = plot_data['Course|Order'].apply(
            lambda x: f"{x} (n={course_counts[x]})"
        )
        # Create and save plot
        # make_python_figure = 0 # Turn on when running in vscode. If being run automatically by qlik, keep off.
        # if make_python_figure:
        #     fig = create_violin_plot(plot_data)
        #     fig.savefig("violin.png", bbox_inches='tight')
        # Calculate the above plot in table form with count and average grades for course order
        agg_result = query.groupby('Course|Order').agg({'Avg Grade': ['count', 'mean']})
        agg_result.columns = ['Count', 'Average Grade']
        agg_result = agg_result.sort_values(by=('Average Grade'),  ascending=False)
        agg_result.sort_values(by=('Count'),  ascending=False, inplace=True)

        agg_result.reset_index(inplace=True)
        print(agg_result)
    return course_counts_initial, agg_result, plot_data,query

if __name__ == "__main__":


    # Get filepath of script and get relative filepaths for log and data folders
    script_path = os.getcwd()
    main_folder_path = os.path.dirname(script_path) # Goes up one folder from script location
    data_path = os.path.join(main_folder_path, 'data')

    try:
        # Step 1: Load the course data
        courses_data = pd.read_parquet(f'{data_path}/Courses/Course_Grade_Data.parquet')
        # Step 2: Handle CLI argument
        # This checks that a csv input file was given
        if len(sys.argv) < 2:
            raise ValueError(r"No input file path provided. Call the function like: python CourseOrder.py 'E:\UBI_Projects\Course Order\data\Selected Courses\ywe4kw_2025-07-08_09-37-48.csv'")

        inFile = sys.argv[1]
        save_file_name = inFile.split('\\')[-1][:-4] # Collect the UserID and date-time for saving the file. 
        output_path = f'{data_path}/Results/{save_file_name}.xlsx'

        # Step 3: Load the Qlik input file
        try:
            qlik_user_input = pd.read_csv(inFile)
        except Exception as e:
            error_message = "Error: Qlik selection file failed to import"
            raise e  # re-raise so outer except catches and prints the right message

        # Step 4: Process course list
        course_list = qlik_user_input['SelectedCourseAndTerm'].fillna('').str.strip().tolist()
        course_list = [c for c in course_list if str(c).strip()]
        # Extract the unique course codes without the year/semester
        courses_without_term = sorted(set(s.split(' - ')[0] for s in course_list))
        courses_without_term_df = pd.DataFrame(courses_without_term, columns=['Course List'])
        course_count = len(courses_without_term)

        # Step 5: Cutoff value
        try:
            min_cutoff_course_number = int(qlik_user_input['CutoffInput'].iloc[0])
        except:
            min_cutoff_course_number = 0 # Default is 0 if no input

        # Step 6: Include repeats
        include_repeats = qlik_user_input['RepeatYesNo'].iloc[0]
        include_repeats = str(include_repeats).strip().lower() == 'yes' # True if yes, otherwise False. Default is False.  

        # Step 7: Main function call
        if (course_count < 2) or (course_count > 5):
            error_message = 'Error: Please select 2-5 courses'
        else:
            try:
                [course_counts, agg_result, plot_data, query] = analyze_course_order(courses_data, course_list, min_cutoff_course_number, include_repeats)
                plot_data_cols_to_keep = ['Student System ID', 'Avg Grade', 'Course|Order w/ Count']
                if agg_result.shape[0] < 1:
                    raise ValueError('Valid inputs, but no students have taken that combination of classes')
                else:
                    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                        plot_data[plot_data_cols_to_keep].to_excel(writer, sheet_name='Plot Data', index=False)
                        agg_result.to_excel(writer, sheet_name='Summary Table', index=False)
                        courses_without_term_df.to_excel(writer, sheet_name='Course List', index=False)
                    error_message = 'No errors!'
            except ValueError as ve:
                raise ve
            except Exception as e:
                error_message = 'Error: Please check that all inputs are valid'

    except ValueError as ve:
        error_message = str(ve)  # This will say: "No input file path provided."
    except Exception as e:
        error_message = 'Error: base course list failed to import'

    # Update log file
    dt = datetime.datetime.now()
    time = datetime.datetime.now().time()
    # Create two logs for each day, one before 12:00 and one for after. Labeled "AM" and "PM"
    if time < datetime.time(12,0):
        log_time_of_day = 'AM'
    else:
        log_time_of_day = 'PM'
    date = datetime.datetime.now().date()
    log_path = os.path.join(main_folder_path, f'logs\\{date}_{log_time_of_day}_log.txt')
    f = open(log_path, "a")
    f.write(f"----------\n")
    f.write(f"{dt} script ran\n")
    # Print if successfully in virtual environment
    if hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix:
        f.write("   Inside a virtual environment %s\n" % (sys.prefix))
    else:
        f.write("   Not in a virtual environment.")
    f.write(f"   {error_message}\n")
    f.write(f"   Request: {save_file_name}\n")
    try:
        f.write(f"   Results:\n{agg_result}\n")
        f.write(f"   Course list: {course_list}\n")
    except:
        f.write(f"   Course list empty or calculation error\n")
    f.write(f"Script finished\n")
    f.close()

    # Convert the string into a DataFrame for one of the excel file tabs    
    error_df = pd.DataFrame({'Error': [error_message]})
    
    # Add Error message to output excel file
        # If the code above errored and did not create a file, make a new one. Otherwise, append the error message to a new sheet
    print(output_path.split('/')[-1].find('scheduler') != -1)

    if not (output_path.split('/')[-1].find('scheduler') != -1): # qlik auto generates some "scheduler" files when it reloads the app. Don't process these. 
        if os.path.exists(output_path):
            mode = 'a' # append
        else:
            # Create empty sheet that will still load into qlik but display nothing       
            plot_columns = ['Student System ID', 'Avg Grade', 'Course|Order w/ Count']
            plot_data = pd.DataFrame(columns=plot_columns)
            agg_result_columns = ['Course|Order', 'Count', 'Average Grade']
            agg_result = pd.DataFrame(columns=agg_result_columns)
            # courses_without_term_df = pd.DataFrame(columns=['Course List'])

            courses_without_term = sorted(set(s.split(' - ')[0] for s in course_list))
            courses_without_term_df = pd.DataFrame(courses_without_term, columns=['Course List'])
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                plot_data.to_excel(writer, sheet_name='Plot Data', index=False)
                agg_result.to_excel(writer, sheet_name='Summary Table', index=False)
                courses_without_term_df.to_excel(writer, sheet_name='Course List', index=False)
            mode = 'w' # write
        with pd.ExcelWriter(f'{data_path}\\Results\\{save_file_name}.xlsx', 
                            mode = 'a', # a = append, to same file that was created above.
                            engine='openpyxl',
                            if_sheet_exists="replace") as writer:
            # Add error message
            error_df.to_excel(writer, sheet_name='Error Message', index=False)

    print(error_message)