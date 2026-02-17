import pandas as pd
import matplotlib.pyplot as plt

def main():
    # 1. Read the dataset file from the repository
    file_path = 'Grad Program Exit Survey Data 2024.xlsx'
    df = pd.read_csv(file_path, encoding='cp1252')

    # 2. Clean and reshape the data
    # Find all columns related to the course ranking question (Q35)
    q35_cols = [col for col in df.columns if col.startswith('Q35')]
    
    # Extract clean course names from Row 0 (the question text row)
    course_names = {}
    for col in q35_cols:
        question_text = df.loc[0, col]
        course_name = question_text.split('- ')[-1].strip()
        course_names[col] = course_name

    # Filter out Qualtrics metadata rows (Row 0 and 1)
    df_data = df.iloc[2:].copy()
    
    # Convert ranking responses to numeric formats for math operations
    for col in q35_cols:
        df_data[col] = pd.to_numeric(df_data[col], errors='coerce')

    # Calculate average ranking (lower score = higher preference)
    mean_ranks = df_data[q35_cols].mean().rename(course_names)

    # 3. Generate a ranking output
    ranked_courses = mean_ranks.sort_values()
    
    # Save the ranking to a CSV file for the grader
    ranked_courses.to_csv('final_ranking.csv', header=['Average_Rank'], index_label='Course')
    print("Ranking successfully saved to final_ranking.csv")

    # 4. Create one figure communicating the ranking
    plt.figure(figsize=(10, 6))
    ranked_courses.plot(kind='barh', color='#4C72B0', edgecolor='black')
    
    plt.xlabel('Average Student Rank (Lower indicates higher preference)')
    plt.ylabel('MAcc Core Courses')
    plt.title('MAcc Core Course Rankings - Most to Least Beneficial (2024)')
    
    # Invert y-axis to put the #1 most preferred course at the very top of the visual
    plt.gca().invert_yaxis() 
    plt.tight_layout()
    
    # Save the figure output
    plt.savefig('course_ranking_figure.png', dpi=300)
    print("Figure successfully saved to course_ranking_figure.png")

if __name__ == "__main__":
    main()
