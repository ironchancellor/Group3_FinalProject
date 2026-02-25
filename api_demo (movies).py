import requests
from openpyxl import Workbook

# ====== 1) CONFIG ======
API_KEY = "e421086260fe93dbfa2406efa0e9cf5f"  # <-- put your key here
BASE_URL = "https://api.themoviedb.org/3/discover/movie"

def fetch_movies_for_year(year, page=1): # Helper function to fetch movies for a given year and page
    params = {
        "api_key": API_KEY,
        "sort_by": "vote_average.desc",   # sort by rating
        "primary_release_year": year,     # filter by year
        "vote_count.gte": 50,             # avoid movies with very few votes
        "page": page
    }
    resp = requests.get(BASE_URL, params=params) # It asks the API for data, gets the data back in JSON format, converts it into Python data, and returns it.
    resp.raise_for_status() # Raises HTTPError for bad responses (4xx/5xx)
    return resp.json() # Converts response JSON into Python objects

def get_top_5_movies_for_year(year): # Processes the raw data to extract only the useful fields, sorts the movies by rating and vote count, and returns the top 5.
    # One page usually enough since sorted by rating; you can loop pages if needed
    data = fetch_movies_for_year(year, page=1) # It calls the helper function to fetch movies for the specified year and page.
    results = data.get("results", []) # It retrieves the list of movies from the JSON data, defaulting to an empty list if not found.

    # Keep only useful fields
    movies = [] # It initializes an empty list to store the processed movie data.
    for m in results: # It iterates over each movie in the results and extracts only the relevant fields (title, release date, rating, vote count, overview) into a new dictionary, which is then appended to the movies list.
        movies.append({
            "title": m.get("title"), # It extracts the title of the movie from the original data.
            "release_date": m.get("release_date"), # It extracts the release date of the movie from the original data.
            "rating": m.get("vote_average"), # It extracts the average vote (rating) of the movie from the original data.
            "vote_count": m.get("vote_count"), # It extracts the total number of votes for the movie from the original data.
            "overview": m.get("overview") # It extracts the overview (summary) of the movie from the original data.
        })

    # Sort again in Python as a safety net, then take top 5
    movies_sorted = sorted( # It sorts the list of movies in Python as a safety net, using a lambda function that sorts primarily by rating (treating None as 0) and secondarily by vote count (also treating None as 0), in descending order.
        movies,
        key=lambda x: (x["rating"] or 0, x["vote_count"] or 0),
        reverse=True
    )
    return movies_sorted[:5] # It returns the top 5 movies from the sorted list.

# ====== 2) WRITE TO EXCEL ======
def write_movies_to_excel(movies, year, filename=None): # It takes the list of movies, the year, and an optional filename. If no filename is provided, it generates one based on the year. It creates a new Excel workbook and worksheet, writes the header row, and then writes each movie's data into subsequent rows. Finally, it saves the workbook to a file.
    if filename is None: # If no filename is provided, it generates a default filename based on the year.
        filename = f"top_5_movies_{year}.xlsx" #  It creates a filename string using an f-string that includes the year.

    wb = Workbook() # It creates a new Excel workbook using the openpyxl library.
    ws = wb.active  # It gets the active worksheet from the workbook, which is where the movie data will be written.
    ws.title = f"Top 5 Movies {year}" # It sets the title of the worksheet to indicate that it contains the top 5 movies for the specified year.

    # Header row
    headers = ["Rank", "Title", "Release Date", "Rating", "Vote Count", "Overview"] # It defines a list of headers for the columns in the Excel sheet, which include Rank, Title, Release Date, Rating, Vote Count, and Overview.
    ws.append(headers) # It appends the header row to the worksheet, which will be the first row in the Excel sheet.

    # Data rows
    for i, m in enumerate(movies, start=1): # It iterates over the list of movies, using enumerate to get both the index (starting at 1 for rank) and the movie data. For each movie, it appends a new row to the worksheet with the rank, title, release date, rating, vote count, and overview.
        ws.append([
            i,
            m["title"],
            m["release_date"],
            m["rating"],
            m["vote_count"],
            m["overview"]
        ]) #  It appends a new row to the worksheet for each movie, containing the rank (i), title, release date, rating, vote count, and overview.

    # Simple “creative twist”: add a short comment/summary row at bottom
    ws.append([]) # It appends an empty row to the worksheet for spacing before adding a summary comment.
    ws.append([f"These are the top 5 rated movies of {year} (min 50 votes)."]) # It appends a new row to the worksheet that contains a summary comment about the data, indicating that these are the top 5 rated movies of the specified year with a minimum of 50 votes.

    wb.save(filename) # It saves the workbook to a file with the specified filename, which will create an Excel file on the disk containing the movie data.
    print(f"Excel file saved as: {filename}") # It prints a message to the console indicating that the Excel file has been saved, along with the filename.

# ====== 3) MAIN FLOW ======
if __name__ == "__main__": # It checks if the script is being run directly (as the main program) rather than imported as a module. If it is the main program, it executes the code within this block.
    year_str = input("Enter a year (e.g., 2023): ").strip() # It prompts the user to enter a year, reads the input as a string, and removes any leading or trailing whitespace.
    if not year_str.isdigit(): # It checks if the input string consists only of digits (i.e., is a valid numeric year). If not, it raises a ValueError indicating that the year must be numeric.
        raise ValueError("Year must be numeric.") # If the input is valid, it converts the string to an integer and stores it in the variable 'year'.
    year = int(year_str) #  It converts the input string to an integer and assigns it to the variable 'year'.
    movies = get_top_5_movies_for_year(year) # It calls the function to get the top 5 movies for the specified year and stores the result in the variable 'movies'.

    if not movies: # It checks if the list of movies is empty. If it is empty, it prints a message indicating that no movies were found for the specified year. Otherwise, it prints the top 5 movies along with their ratings and calls the function to write the movies to an Excel file.
        print(f"No movies found for year {year}.") #  If the list of movies is empty, it prints a message indicating that no movies were found for the specified year.
    else:
        print(f"Top 5 movies for {year}:") # If the list of movies is not empty, it prints a header indicating that these are the top 5 movies for the specified year.
        for i, m in enumerate(movies, start=1): # It iterates over the list of movies, using enumerate to get both the index (starting at 1 for rank) and the movie data. For each movie, it prints the rank, title, and rating to the console.
            print(f"{i}. {m['title']} ({m['rating']})") # It prints the rank (i), title, and rating of each movie in a formatted string.

        write_movies_to_excel(movies, year) # It calls the function to write the list of movies to an Excel file, passing the movies and the year as arguments.
