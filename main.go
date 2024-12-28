package main

import (
	_ "embed"
	"encoding/csv"
	"fmt"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

var (
  size         int     = 1
	orientation  string  = "landscape"
	BottomMargin float64 = 0.15
	TopMargin    float64 = 0.25
	LeftMargin   float64 = 0.35
	RightMargin  float64 = 0.2
	bgFill 			 string  = "#002060"
)
const dateFormat string = "1/2/2006"
const timeFormat string = "15:04"
const outputTimeFormat string = "3:04 PM"
const outputCsvFolder string = "outputDataCsv"
const outputExcelFolder string = "outputDataExcel"
 
var data []map[string]interface{}
	
func main() {
	 if _, err := os.Stat(outputCsvFolder); os.IsNotExist(err) {
        os.Mkdir(outputCsvFolder, os.ModePerm)
		if _, err := os.Stat(outputExcelFolder); os.IsNotExist(err) {
				os.Mkdir(outputExcelFolder, os.ModePerm)
    }
	}
	
	// Read and parse CSV files
    files := []string{"data/7U.csv", "data/10U.csv", "data/12U.csv", "data/15U.csv"}
    for _, file := range files {
        if err := readCSV(file); err != nil {
            fmt.Println(err)
            return
        }
    }

   sortDataByDateTimeAndLocation(data)

     // Write sorted data to separate CSV files by date
	if err := writeCSVByDate(); err != nil {
		fmt.Println(err)
		return
	}
	// Process each CSV in the output folder
   dirEntries, err := os.ReadDir(outputCsvFolder)
if err != nil {
    fmt.Println("Error reading directory:", err)
    return
}

for _, entry := range dirEntries {
    if entry.IsDir() || !strings.HasSuffix(entry.Name(), ".csv") {
        continue
    }
    csvPath := filepath.Join(outputCsvFolder, entry.Name())
    fmt.Println("Processing:", csvPath)

        // Read current CSV rows
        rows, err := readCSVFile(csvPath)
        if err != nil {
            fmt.Println("Error reading CSV:", err)
            continue
        }

        // Insert missing field rows
        updatedRows := fillMissingFields(rows)

        // Write back to CSV
        if err := writeUpdatedCSV(csvPath, updatedRows); err != nil {
            fmt.Println("Error writing CSV:", err)
            continue
        }
    };
		processUpdatedCSVs()

}

// After writing updated CSVs, read them again and call writeExcel on each.
func processUpdatedCSVs() error {
    dirEntries, err := os.ReadDir(outputCsvFolder)
    if err != nil {
        return err
    }

    for _, entry := range dirEntries {
        if entry.IsDir() || !strings.HasSuffix(entry.Name(), ".csv") {
            continue
        }

        csvPath := filepath.Join(outputCsvFolder, entry.Name())
        fmt.Println("Reading updated CSV:", csvPath)

				// Read rows from updated CSV
        rows, err := readCSVFile(csvPath)
        if err != nil {
            fmt.Println("Error reading CSV:", err)
            continue
        }

				   // Convert [][]string to []map[string]interface{}
        dataMaps := rowsToDataMaps(rows)

       
        // Write Excel file with matching name
        excelFilename := filepath.Join(
            outputExcelFolder,
            strings.TrimSuffix(entry.Name(), ".csv")+".xlsx",
        )
        if err := writeExcel(excelFilename, dataMaps); err != nil {
            fmt.Println("Error writing Excel:", err)
            continue
        }
    }
    return nil
}

// Convert the [][]string (header + data rows) to []map[string]interface{}.
// Assumes columns: Home, Away, Date, Time, Location
func rowsToDataMaps(rows [][]string) []map[string]interface{} {
    if len(rows) < 2 {
        return nil
    }
    header := rows[0]
    var out []map[string]interface{}

    for _, r := range rows[1:] {
        if len(r) < len(header) {
            continue
        }
        rowMap := map[string]interface{}{
            header[0]: r[0],
            header[1]: r[1],
            header[2]: r[2],
            header[3]: r[3],
            header[4]: r[4],
        }
        out = append(out, rowMap)
    }
    return out
}


// readCSVFile reads all rows from a CSV.
func readCSVFile(csvPath string) ([][]string, error) {
    f, err := os.Open(csvPath)
    if err != nil {
        return nil, err
    }
    defer f.Close()

    reader := csv.NewReader(f)
    var rows [][]string
    for {
        record, err := reader.Read()
        if err == io.EOF {
            break
        }
        if err != nil {
            return nil, err
        }
        rows = append(rows, record)
    }
    return rows, nil
}

func fillMissingFields(rows [][]string) [][]string {
    if len(rows) < 2 {
        return rows
    }

    header := rows[0]
    dataRows := rows[1:]
    timeMap := make(map[string][][]string)

    // Group rows by Time (column index 3)
    for _, r := range dataRows {
        if len(r) < 5 {
            continue
        }
        t := r[3]
        timeMap[t] = append(timeMap[t], r)
    }

    var result [][]string
    result = append(result, header)

    // For each time group, ensure Field #1..#6 are present
    for t, groupRows := range timeMap {
        if len(groupRows) == 0 || len(groupRows[0]) < 5 {
            continue
        }
        dateVal := groupRows[0][2]

        fieldsPresent := make(map[int]bool)
        fieldToRow := make(map[int][]string)

        // Identify fields present
        for _, gr := range groupRows {
            parts := strings.Split(gr[4], "#")
            if len(parts) == 2 {
                var fNum int
                fmt.Sscanf(strings.TrimSpace(parts[1]), "%d", &fNum)
                fieldsPresent[fNum] = true
                fieldToRow[fNum] = gr
            }
        }

        // Insert rows for missing fields
        for fNum := 1; fNum <= 6; fNum++ {
            if fieldsPresent[fNum] {
                result = append(result, fieldToRow[fNum])
            } else {
                filler := []string{"Open Field", "Open Field", dateVal, t, fmt.Sprintf("Field #%d", fNum)}
                result = append(result, filler)
            }
        }
				// Add empty row after each time block (after Field #6)
        emptyRow := []string{"", "", "", "", ""}
        result = append(result, emptyRow)
    }
    return result
}


// writeUpdatedCSV overwrites the file with updated rows.
func writeUpdatedCSV(csvPath string, rows [][]string) error {
    f, err := os.Create(csvPath)
    if err != nil {
        return err
    }
    defer f.Close()

    writer := csv.NewWriter(f)
    defer writer.Flush()

    for _, r := range rows {
        if err := writer.Write(r); err != nil {
            return err
        }
    }
    return nil
}


func sortDataByDateTimeAndLocation(data []map[string]interface{}) {
    sort.Slice(data, func(i, j int) bool {
        date1, _ := time.Parse(dateFormat, data[i]["Date"].(string))
        date2, _ := time.Parse(dateFormat, data[j]["Date"].(string))
        if date1.Equal(date2) {
            time1, _ := time.Parse(timeFormat, data[i]["Time"].(string))
            time2, _ := time.Parse(timeFormat, data[j]["Time"].(string))
            if time1.Equal(time2) {
                loc1 := strings.Split(data[i]["Location"].(string), "#")
                loc2 := strings.Split(data[j]["Location"].(string), "#")
                if len(loc1) < 2 || len(loc2) < 2 {
                    // Fallback compare if no "#"
                    return data[i]["Location"].(string) < data[j]["Location"].(string)
                }
                field1, _ := strconv.Atoi(loc1[1])
                field2, _ := strconv.Atoi(loc2[1])
                return field1 < field2
            }
            return time1.Before(time2)
        }
        return date1.Before(date2)
    })
}


func readCSV(filename string) error {
    file, err := os.Open(filename)
    if err != nil {
        return err
    }
    defer file.Close()

    reader := csv.NewReader(file)
    records, err := reader.ReadAll()
    if err != nil {
        return err
    }

		  fileName := strings.TrimSuffix(strings.TrimPrefix(filename, "data/"), ".csv")

    // Skip header row
    for _, record := range records[1:] {
        data = append(data, map[string]interface{}{
              "Home":     fmt.Sprintf("%s %s", fileName, record[0]),
            "Away":     fmt.Sprintf("%s %s", fileName, record[1]),
            "Date":     record[2],
            "Time":     record[3],
            "Location": record[4],
        })
    }
    return nil
}

func writeCSVByDate() error {
    // Group data by date
    dateGroups := make(map[string][]map[string]interface{})
    for _, row := range data {
        date := row["Date"].(string)
        dateGroups[date] = append(dateGroups[date], row)
    }
    // Write each group to a separate CSV file
    for date, rows := range dateGroups {
		filename := fmt.Sprintf("%s/sorted_schedule_%s.csv", outputCsvFolder, strings.ReplaceAll(date, "/", "-"))
        if err := writeCSV(filename, rows); err != nil {
          fmt.Println(err)  
					return err
        }
    }
    return nil
}

func writeCSV(filename string, rows []map[string]interface{}) error {
    file, err := os.Create(filename)
    if err != nil {
        return err
    }
    defer file.Close()

    writer := csv.NewWriter(file)
    defer writer.Flush()

    // Write headers
    headers := []string{"Home", "Away", "Date", "Time", "Location"}
    if err := writer.Write(headers); err != nil {
        return err
    }

    // Write data
    for _, row := range rows {
			 // Parse and format the time
        parsedTime, _ := time.Parse(timeFormat, row["Time"].(string))
        formattedTime := parsedTime.Format(outputTimeFormat)
        record := []string{
            row["Home"].(string),
            row["Away"].(string),
            row["Date"].(string),
            formattedTime,
            row["Location"].(string),
        }
        if err := writer.Write(record); err != nil {
            return err
        }
    }

    return nil
}


// getWeekNumber returns the week number for a given date, starting on Monday
func getWeekNumber(startDate time.Time, daysToAdd int) int {
    newDate := startDate.AddDate(0, 0, daysToAdd)
    _, week := newDate.ISOWeek()
    return week
}

// calculateWeekNumber takes start date, end date, and current date, and returns the week number
func calculateWeekNumber(startDate, endDate, currentDate time.Time) (int, error) {
    if currentDate.Before(startDate) || currentDate.After(endDate) {
        return 0, fmt.Errorf("current date is out of range")
    }

    daysSinceStart := int(currentDate.Sub(startDate).Hours() / 24)
    weekNumber := getWeekNumber(startDate, daysSinceStart)
    return weekNumber, nil
}






func writeExcel(filename string, data []map[string]interface{}) error {
	f := excelize.NewFile()

	defer func() {
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()

	if err := f.SetDefaultFont("Arial Rounded MT Bold"); err != nil {
		fmt.Println(err)
	}

	if err := f.SetPageLayout("Sheet1", &excelize.PageLayoutOptions{
		Size:        &size,
		Orientation: &orientation,
	}); err != nil {
		fmt.Println(err)
	}

	if err := f.SetPageMargins("Sheet1", &excelize.PageLayoutMarginsOptions{
		Left:   &LeftMargin,
		Right:  &RightMargin,
		Top:    &TopMargin,
		Bottom: &BottomMargin,
	}); err != nil {
		fmt.Println(err)
	}
	if err := f.MergeCell("Sheet1", "B1", "B3"); err != nil {
		fmt.Println(err)
	}
	if err := f.MergeCell("Sheet1", "E1", "F3"); err != nil {
		fmt.Println(err)
	}
	 enable, disable := true, false

	if err := f.AddPicture("Sheet1", "B1",  "./images/FlagLogo_x0.5.png", &excelize.GraphicOptions{
			PrintObject:     	&enable,
			Locked:          	&disable,
			OffsetX:         	5,
			OffsetY:         	5,
			ScaleX: 					0.3,
			ScaleY: 					0.5,
			AutoFit: 			 		false,
			LockAspectRatio: 	true,
			Positioning: 			"absolute",
			
		}); err != nil {
        fmt.Println(err)
        return err
    }

		if err := f.AddPicture("Sheet1", "E1", "./images/PlayFootball@0.25x.png", &excelize.GraphicOptions{
						PrintObject:     	&enable,
						Locked:          	&disable,
            OffsetX:         	50,
            OffsetY:         	10,
						ScaleX: 					0.4,
						ScaleY: 					0.4,
						AutoFit: 			 		false,
						LockAspectRatio: 	true,
						Positioning: 			"absolute",
						
        },); err != nil {
        fmt.Println(err)
        return err
    }		

		dateStr := data[2]["Date"].(string)
		// Example dates
    inputStartDate := "1/3/2025"
    inputEndDate := "3/24/2025"
    currentDate, err := time.Parse(dateFormat, dateStr)
		if err != nil {
    fmt.Println("Error parsing date data from table:", err)
    return err
} 

	startDate, err := time.Parse(dateFormat, inputStartDate)
    if err != nil {
        fmt.Println("Error parsing start date:", err)
        return err
    }

	endDate, err := time.Parse(dateFormat, inputEndDate)
    if err != nil {
        fmt.Println("Error parsing end date:", err)
        return err
    }

    weekNumber, err := calculateWeekNumber(startDate, endDate, currentDate )
    if err != nil {
        fmt.Println(err)
        return err
    }

    fmt.Printf("Current week number: %d\n", weekNumber)

	var rowTitle = []string{"HOME", "AWAY", "DATE", "TIME", "LOCATION"}
	sheet := "Sheet1" 
	getDate, _ := time.Parse(dateFormat, dateStr)
	week := fmt.Sprintf("Week #%d", weekNumber)
	location := "Englewood, Florida"
	f.SetCellValue(sheet, "C1", week)
	f.SetCellValue(sheet, "C2", getDate.Format("January 2, 2006"))
	f.SetCellValue(sheet, "C3", location)
	f.SetCellValue(sheet, "B4", rowTitle[0])
	f.SetCellValue(sheet, "C4", rowTitle[1])
	f.SetCellValue(sheet, "D4", rowTitle[2])
	f.SetCellValue(sheet, "E4", rowTitle[3])
	f.SetCellValue(sheet, "F4", rowTitle[4])
	// Replace with the name of your sheet
	startRow := 5  
	for i, row := range data {
		rowNumber := startRow + i
		// Replace placeholders or write to specific cells
		f.SetCellValue(sheet, fmt.Sprintf("B%d", rowNumber), row["Home"])
		f.SetCellValue(sheet, fmt.Sprintf("C%d", rowNumber), row["Away"])
		f.SetCellValue(sheet, fmt.Sprintf("D%d", rowNumber), row["Date"])
		f.SetCellValue(sheet, fmt.Sprintf("E%d", rowNumber), row["Time"])
		f.SetCellValue(sheet, fmt.Sprintf("F%d", rowNumber), row["Location"])
	}

	if err := f.SetColWidth("Sheet1", "A", "A", 0.8); err != nil {
		fmt.Println(err)
	}
		if err := f.SetColWidth("Sheet1", "B", "C", 34); err != nil {
		fmt.Println(err)
	}
		if err := f.SetColWidth("Sheet1", "D",  "D",15); err != nil {
		fmt.Println(err)
	}
	if err := f.SetColWidth("Sheet1", "E",  "E",12); err != nil {
		fmt.Println(err)
	}
	if err := f.SetColWidth("Sheet1", "F",  "F",15); err != nil {
		fmt.Println(err)
	}
		if err := f.SetColWidth("Sheet1", "G",  "G",0.8); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 1, 28 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 2, 28 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 3, 24 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 4, 24 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 5, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 6, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 7, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 8, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 9, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 10, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 11, 5 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 12, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 13, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 14, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 15, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 16, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 17, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 18, 5 ); err != nil {
		fmt.Println(err)
	}
		if err := f.SetRowHeight("Sheet1", 19, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 20, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 21, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 22, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 23, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 24, 26 ); err != nil {
		fmt.Println(err)
	}
	if err := f.SetRowHeight("Sheet1", 25, 5 ); err != nil {
		fmt.Println(err)
	}
	bgStyle, err := f.NewStyle(&excelize.Style{Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{bgFill}}})
	if err != nil {
		fmt.Println(err)
	}
	if err := f.SetCellStyle("Sheet1", "B1", "F4", bgStyle); err != nil {	
		fmt.Println(err)
	}
		if err := f.SetCellStyle("Sheet1", "A1", "A25", bgStyle); err != nil {	
		fmt.Println(err)
	}
		if err := f.SetCellStyle("Sheet1", "G1", "G25", bgStyle); err != nil {	
		fmt.Println(err)
	}
	if err := f.SetCellStyle("Sheet1", "B11", "F11", bgStyle); err != nil {	
		fmt.Println(err)
	}
	if err := f.SetCellStyle("Sheet1", "B18", "F18", bgStyle); err != nil {	
		fmt.Println(err)
	}
	if err := f.SetCellStyle("Sheet1", "B25", "F25", bgStyle); err != nil {	
		fmt.Println(err)
	}
	allBorder, err := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical: "center",
		},
		Font: &excelize.Font{
			Size:   14,
			Color:  "000000",
		},
	})
	if err != nil {
		fmt.Println(err)
	}
	if err := f.SetCellStyle("Sheet1", "B5", "F10", allBorder); err != nil {
		fmt.Println(err)
	}
		if err := f.SetCellStyle("Sheet1", "B12", "F17", allBorder); err != nil {
		fmt.Println(err)
	}
		if err := f.SetCellStyle("Sheet1", "B19", "F24", allBorder); err != nil {
		fmt.Println(err)
	}
	weekStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{bgFill}},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical: "top",
		},
		Font: &excelize.Font{
			Bold:   true, 
			Size:   26,
			Color:  "#FFFFFF",
		},
	})
	if err != nil {
		fmt.Println(err)
	}
		if err := f.SetCellStyle("Sheet1", "C1", "C1", weekStyle); err != nil {
		fmt.Println(err)
	}

		dateStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{bgFill}},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
		Font: &excelize.Font{
			Size:   16,
			Color:  "#FFFFFF",
		},
	})
	if err != nil {
		fmt.Println(err)
	}
		if err := f.SetCellStyle("Sheet1", "C2", "C2", dateStyle); err != nil {
		fmt.Println(err)
	}
			locationStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{bgFill}},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical: "center",
		},
		Font: &excelize.Font{
			Size:   12,
			Color:  "#FFFFFF",
		},
	})
	if err != nil {
		fmt.Println(err)
	}
		if err := f.SetCellStyle("Sheet1", "C3", "C3", locationStyle); err != nil {
		fmt.Println(err)
	}
		rowTitleStyle, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{bgFill}},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical: "center",
		},
		Font: &excelize.Font{
			Size:   12,
			Color:  "#FFFFFF",
		},
	})
	if err != nil {
		fmt.Println(err)
	}
		if err := f.SetCellStyle("Sheet1", "B4", "F4", rowTitleStyle); err != nil {
		fmt.Println(err)
	}


		 // Set value of a cell.
    f.SetCellValue("Week1", "B2", 100)
    // Set active sheet of the workbook.
    // Save spreadsheet by the given path.
    if err := f.SaveAs(filename); err != nil {
        fmt.Println(err)
    }
		return nil
}