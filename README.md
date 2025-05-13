This notebook analyzes the impact of course-taking order on student outcomes (e.g., GPA) at the University of Virginia. It was created in response to a question about whether students perform differently depending on the sequence in which they take certain key classes—e.g., is it better to take Calculus before Linear Algebra, or the other way around?

The code is a query-style script with plans to integrate into UVA's UBI (university business intelligence). The user inputs a list of courses to analyze, and the script outputs a violin plot which visualizes how students have historically performed in those courses with respect to the order that a course is taken.

In the violin plot, each course is followed by a number (1, 2, or 3 for example). The number corresponds to the order in the sequence in which the student took a course. If a student has two courses with the same number, that means they took those courses in the same semester.
