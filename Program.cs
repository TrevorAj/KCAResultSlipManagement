using System;
using MySql.Data.MySqlClient;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Linq;
using System.Diagnostics;

class Program
{
    static string connectionString = "Server=localhost;Database=KCA_Results;Uid=root;Pwd=Root;";

    static void Main()
    {
        // Open the logo image at startup
        string logoPath = @"C:\Users\Dell\source\New folder\LOGO.png";

        if (File.Exists(logoPath))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = logoPath,
                UseShellExecute = true
            });
        }

        while (true)
        {
            Console.Clear();
            Console.WriteLine("Welcome to the KCA Student Result Management System");
            Console.WriteLine("1. Enter Student Marks");
            Console.WriteLine("2. Generate Result Slip");
            Console.WriteLine("3. View Students with Missing Marks");
            Console.WriteLine("4. Add New Unit");
            Console.WriteLine("5. Exit");
            Console.Write("Choose an option: ");

            if (int.TryParse(Console.ReadLine(), out int choice))
            {
                switch (choice)
                {
                    case 1:
                        EnterStudentMarks();
                        break;
                    case 2:
                        GenerateResultSlip();
                        break;
                    case 3:
                        ViewMissingMarks();
                        break;
                    case 4:
                        AddNewUnit();
                        break;
                    case 5:
                        return;
                    default:
                        Console.WriteLine("Invalid choice.");
                        break;
                }
            }

            Console.WriteLine("\nPress any key to continue...");
            Console.ReadKey();
        }
    }


    static void EnterStudentMarks()
    {
        Console.Write("Enter Admission Number: ");
        string admissionNumber = Console.ReadLine() ?? "";

        using (MySqlConnection conn = new MySqlConnection(connectionString))
        {
            try
            {
                conn.Open();
                
                // Verify student exists
                if (!VerifyStudent(conn, admissionNumber))
                {
                    Console.Write("Student not found! Would you like to add them? (yes/no): ");
                    if ((Console.ReadLine() ?? "").ToLower() == "yes")
                    {
                        AddStudent(admissionNumber);
                    }
                    else
                    {
                        return;
                    }
                }

                // Display available units
                Console.WriteLine("\nAvailable Units:");
                string unitsQuery = "SELECT unit_code, unit_name FROM units";
                using (MySqlCommand cmd = new MySqlCommand(unitsQuery, conn))
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    if (!reader.HasRows)
                    {
                        Console.WriteLine("No units found in the system!");
                        return;
                    }
                    while (reader.Read())
                    {
                        Console.WriteLine($"{reader["unit_code"]} - {reader["unit_name"]}");
                    }
                }

                Console.Write("\nEnter Unit Code: ");
                string unitCode = Console.ReadLine()?.Trim().ToUpper() ?? "";

                // Verify unit exists
                if (!VerifyUnit(conn, unitCode))
                {
                    Console.WriteLine("Unit not found!");
                    return;
                }

                // Check if marks already exist
                string checkQuery = "SELECT COUNT(*) FROM results WHERE admission_number = @admission AND unit_code = @unit";
                using (MySqlCommand checkCmd = new MySqlCommand(checkQuery, conn))
                {
                    checkCmd.Parameters.AddWithValue("@admission", admissionNumber);
                    checkCmd.Parameters.AddWithValue("@unit", unitCode);
                    if (Convert.ToInt32(checkCmd.ExecuteScalar()) > 0)
                    {
                        Console.WriteLine("Marks already exist for this unit. Please use update function (not yet implemented).");
                        return;
                    }
                }

                // Input marks
                int?[] marks = new int?[5];
                string[] markTypes = { "Assignment 1", "Assignment 2", "Assignment 3", "CAT 1", "CAT 2" };

                for (int i = 0; i < marks.Length; i++)
                {
                    bool validInput = false;
                    while (!validInput)
                    {
                        Console.Write($"Enter marks for {markTypes[i]} (0-10) or press Enter to skip: ");
                        string input = Console.ReadLine() ?? "";
                        
                        if (string.IsNullOrEmpty(input))
                        {
                            marks[i] = null;
                            validInput = true;
                        }
                        else if (int.TryParse(input, out int mark) && mark >= 0 && mark <= 10)
                        {
                            marks[i] = mark;
                            validInput = true;
                        }
                        else
                        {
                            Console.WriteLine("Invalid input! Please enter a number between 0 and 10.");
                        }
                    }
                }

                // Input exam marks
                int? exam = null;
                bool validExamInput = false;
                while (!validExamInput)
                {
                    Console.Write("Enter Exam marks (0-50) or press Enter to skip: ");
                    string examInput = Console.ReadLine() ?? "";
                    
                    if (string.IsNullOrEmpty(examInput))
                    {
                        exam = null;
                        validExamInput = true;
                    }
                    else if (int.TryParse(examInput, out int examMark) && examMark >= 0 && examMark <= 50)
                    {
                        exam = examMark;
                        validExamInput = true;
                    }
                    else
                    {
                        Console.WriteLine("Invalid input! Please enter a number between 0 and 50.");
                    }
                }

                // Calculate total and grade
                int? total = marks.Sum(m => m ?? 0) + (exam ?? 0);
                string grade = exam.HasValue && marks.All(m => m.HasValue) ? GetGrade(total.Value) : "IP";

                // Insert into database
                string insertQuery = @"INSERT INTO results 
                    (admission_number, unit_code, assignment1, assignment2, assignment3, cat1, cat2, exam, total, grade) 
                    VALUES (@admission, @code, @a1, @a2, @a3, @c1, @c2, @exam, @total, @grade)";

                using (MySqlCommand cmd = new MySqlCommand(insertQuery, conn))
                {
                    cmd.Parameters.AddWithValue("@admission", admissionNumber);
                    cmd.Parameters.AddWithValue("@code", unitCode);
                    cmd.Parameters.AddWithValue("@a1", (object?)marks[0] ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@a2", (object?)marks[1] ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@a3", (object?)marks[2] ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@c1", (object?)marks[3] ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@c2", (object?)marks[4] ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@exam", (object?)exam ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@total", (object?)total ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@grade", grade);

                    cmd.ExecuteNonQuery();
                    Console.WriteLine("\nMarks entered successfully!");
                    
                    // Show summary
                    Console.WriteLine("\nMarks Summary:");
                    Console.WriteLine($"Student: {admissionNumber}");
                    Console.WriteLine($"Unit: {unitCode}");
                    for (int i = 0; i < marks.Length; i++)
                    {
                        Console.WriteLine($"{markTypes[i]}: {marks[i]?.ToString() ?? "Not entered"}");
                    }
                    Console.WriteLine($"Exam: {exam?.ToString() ?? "Not entered"}");
                    Console.WriteLine($"Total: {total}");
                    Console.WriteLine($"Grade: {grade}");
                }
            }
            catch (MySqlException ex)
            {
                Console.WriteLine($"Database error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    static void GenerateResultSlip()

    {
        try
        {
            Console.Write("Enter Admission Number: ");
            string admissionNumber = Console.ReadLine() ?? "";

            // Create a directory for the PDFs if it doesn't exist
            string pdfDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "KCAResultSlips"
            );
            Directory.CreateDirectory(pdfDirectory);

            // Create the PDF file path with timestamp to ensure uniqueness
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"ResultSlip_{admissionNumber.Replace("/", "_")}_{timestamp}.pdf";
            string filePath = Path.Combine(pdfDirectory, fileName);

            using (MySqlConnection conn = new MySqlConnection(connectionString))
            {
                conn.Open();
                if (!VerifyStudent(conn, admissionNumber))
                {
                    Console.WriteLine("Student not found!");
                    return;
                }

                // Get student details
                string studentQuery = "SELECT * FROM students WHERE admission_number = @admission";
                string studentName = "";
                string programme = "";

                using (MySqlCommand cmd = new MySqlCommand(studentQuery, conn))
                {
                    cmd.Parameters.AddWithValue("@admission", admissionNumber);
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            studentName = reader["name"].ToString() ?? "";
                            programme = reader["programme"].ToString() ?? "";
                        }
                    }
                }

                // Create PDF in memory first
                MemoryStream ms = new MemoryStream();
                Document doc = new Document();
                PdfWriter writer = PdfWriter.GetInstance(doc, ms);
                doc.Open();

                string logoPath = @"C:\Users\Dell\source\New folder\LOGO.png";
                if (File.Exists(logoPath))
                {
                    Image logo = Image.GetInstance(logoPath);
                    logo.ScaleToFit(100f, 100f); // Adjust size
                    logo.Alignment = Element.ALIGN_CENTER;
                    doc.Add(logo);
                }


                // Add content to PDF
                doc.Add(new Paragraph("KCA UNIVERSITY"));
                doc.Add(new Paragraph($"Student: {studentName}"));
                doc.Add(new Paragraph($"Admission Number: {admissionNumber}"));
                doc.Add(new Paragraph($"Programme: {programme}\n\n"));

                PdfPTable table = new PdfPTable(4);
                table.WidthPercentage = 100;

                string[] headers = { "Unit Code", "Unit Name", "Grade", "Status" };
                foreach (string header in headers)
                {
                    PdfPCell cell = new PdfPCell(new Phrase(header));
                    cell.BackgroundColor = new BaseColor(211, 211, 211);
                    table.AddCell(cell);
                }

                string query = @"
                    SELECT r.*, u.unit_name 
                    FROM results r 
                    JOIN units u ON r.unit_code = u.unit_code 
                    WHERE r.admission_number = @admission";

                using (MySqlCommand cmd = new MySqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@admission", admissionNumber);
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        int unitCount = 0;
                        while (reader.Read() && unitCount < 7)
                        {
                            table.AddCell(reader["unit_code"].ToString() ?? "");
                            table.AddCell(reader["unit_name"].ToString() ?? "");
                            table.AddCell(reader["grade"].ToString() ?? "");

                            bool hasMissingMarks = reader["assignment1"] == DBNull.Value ||
                                                 reader["assignment2"] == DBNull.Value ||
                                                 reader["assignment3"] == DBNull.Value ||
                                                 reader["cat1"] == DBNull.Value ||
                                                 reader["cat2"] == DBNull.Value ||
                                                 reader["exam"] == DBNull.Value;

                            table.AddCell(hasMissingMarks ? "Missing Marks" : "Complete");
                            unitCount++;
                        }

                        // Fill remaining slots
                        for (int i = unitCount; i < 7; i++)
                        {
                            for (int j = 0; j < 4; j++)
                            {
                                table.AddCell("-");
                            }
                        }
                    }
                }

                doc.Add(table);
                doc.Close();

                // Write the PDF to file
                byte[] bytes = ms.ToArray();
                File.WriteAllBytes(filePath, bytes);

                Console.WriteLine($"\nPDF generated successfully!");
                Console.WriteLine($"Location: {filePath}");

                // Try to open the PDF
                try
                {
                    new System.Diagnostics.Process
                    {
                        StartInfo = new System.Diagnostics.ProcessStartInfo(filePath)
                        {
                            UseShellExecute = true
                        }
                    }.Start();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\nCould not automatically open the PDF: {ex.Message}");
                    Console.WriteLine($"Please manually open the file at: {filePath}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\nError generating PDF: {ex.Message}");
            Console.WriteLine("Please ensure you have write permissions to the Documents folder.");
        }
    }

    static void ViewMissingMarks()
    {
        Console.Write("Enter Unit Code: ");
        string unitCode = Console.ReadLine() ?? "";

        using (MySqlConnection conn = new MySqlConnection(connectionString))
        {
            conn.Open();
            if (!VerifyUnit(conn, unitCode))
            {
                Console.WriteLine("Unit not found!");
                return;
            }

            string query = @"
                SELECT s.admission_number, s.name, r.* 
                FROM students s 
                JOIN results r ON s.admission_number = r.admission_number 
                WHERE r.unit_code = @unitCode 
                AND (r.assignment1 IS NULL OR r.assignment2 IS NULL OR r.assignment3 IS NULL 
                    OR r.cat1 IS NULL OR r.cat2 IS NULL OR r.exam IS NULL)";

            using (MySqlCommand cmd = new MySqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@unitCode", unitCode);
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    bool hasResults = false;
                    Console.WriteLine("\nStudents with Missing Marks:");
                    Console.WriteLine("----------------------------");

                    while (reader.Read())
                    {
                        hasResults = true;
                        Console.WriteLine($"\nStudent: {reader["name"]} ({reader["admission_number"]})");
                        if (reader["assignment1"] == DBNull.Value) Console.WriteLine("- Missing Assignment 1");
                        if (reader["assignment2"] == DBNull.Value) Console.WriteLine("- Missing Assignment 2");
                        if (reader["assignment3"] == DBNull.Value) Console.WriteLine("- Missing Assignment 3");
                        if (reader["cat1"] == DBNull.Value) Console.WriteLine("- Missing CAT 1");
                        if (reader["cat2"] == DBNull.Value) Console.WriteLine("- Missing CAT 2");
                        if (reader["exam"] == DBNull.Value) Console.WriteLine("- Missing Exam");
                    }

                    if (!hasResults)
                    {
                        Console.WriteLine("No students found with missing marks for this unit.");
                    }
                }
            }
        }
    }

    static void AddNewUnit()
    {
        Console.Write("Enter Unit Code: ");
        string unitCode = Console.ReadLine() ?? "";

        Console.Write("Enter Unit Name: ");
        string unitName = Console.ReadLine() ?? "";

        Console.Write("Enter Faculty: ");
        string faculty = Console.ReadLine() ?? "";

        using (MySqlConnection conn = new MySqlConnection(connectionString))
        {
            conn.Open();
            string query = "INSERT INTO units (unit_code, unit_name, faculty) VALUES (@code, @name, @faculty)";

            using (MySqlCommand cmd = new MySqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@code", unitCode);
                cmd.Parameters.AddWithValue("@name", unitName);
                cmd.Parameters.AddWithValue("@faculty", faculty);

                try
                {
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("Unit added successfully!");
                }
                catch (MySqlException ex)
                {
                    Console.WriteLine($"Error adding unit: {ex.Message}");
                }
            }
        }
    }

    static void AddStudent(string admissionNumber)
    {
        Console.Write("Enter Student Name: ");
        string studentName = Console.ReadLine() ?? "";

        Console.Write("Enter Faculty: ");
        string faculty = Console.ReadLine() ?? "";

        Console.Write("Enter Programme: ");
        string programme = Console.ReadLine() ?? "";

        using (MySqlConnection conn = new MySqlConnection(connectionString))
        {
            conn.Open();
            string query = "INSERT INTO students (admission_number, name, faculty, programme) VALUES (@admission, @name, @faculty, @programme)";

            using (MySqlCommand cmd = new MySqlCommand(query, conn))
            {
                cmd.Parameters.AddWithValue("@admission", admissionNumber);
                cmd.Parameters.AddWithValue("@name", studentName);
                cmd.Parameters.AddWithValue("@faculty", faculty);
                cmd.Parameters.AddWithValue("@programme", programme);

                try
                {
                    cmd.ExecuteNonQuery();
                    Console.WriteLine("Student added successfully!");
                }
                catch (MySqlException ex)
                {
                    Console.WriteLine($"Error adding student: {ex.Message}");
                }
            }
        }
    }

    static bool VerifyStudent(MySqlConnection conn, string admissionNumber)
    {
        string query = "SELECT COUNT(*) FROM students WHERE admission_number = @admission";
        using (MySqlCommand cmd = new MySqlCommand(query, conn))
        {
            cmd.Parameters.AddWithValue("@admission", admissionNumber);
            return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
        }
    }

    static bool VerifyUnit(MySqlConnection conn, string unitCode)
    {
        string query = "SELECT COUNT(*) FROM units WHERE unit_code = @code";
        using (MySqlCommand cmd = new MySqlCommand(query, conn))
        {
            cmd.Parameters.AddWithValue("@code", unitCode);
            return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
        }
    }

    static string GetGrade(int total)
    {
        if (total >= 74) return "A";
        if (total >= 70) return "A-";
        if (total >= 67) return "B+";
        if (total >= 64) return "B";
        if (total >= 60) return "B-";
        if (total >= 57) return "C+";
        if (total >= 54) return "C";
        if (total >= 50) return "C-";
        if (total >= 47) return "D+";
        if (total >= 44) return "D";
        if (total >= 40) return "D-";
        return "F";
    }
}

