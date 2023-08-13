package Assignment1;  //Package name.

//imported libraries in the code to perform.
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.time.Duration;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;

public class Assignment1 {
    public static void main(String[] args) throws IOException {
        System.setProperty("webdriver.chrome.driver", "C:/Users/User/Desktop/Assignment1_Java-main/chromedriver.exe");  //Web driver address.
        WebDriver driver = new ChromeDriver();
        //ChromeOptions chromeOptions = new ChromeOptions();  //Another method to maximize the browser but not used here.
        driver.manage().window().maximize();  //Maximizing the web browser.
        Duration timeout = Duration.ofSeconds(10);  //Delaying for open the browser.
        WebDriverWait wait = new WebDriverWait(driver, timeout);

        // Get current day of the week (e.g., Monday, Tuesday, etc.)
        SimpleDateFormat dateFormat = new SimpleDateFormat("EEEE");
        String currentDay = dateFormat.format(Calendar.getInstance().getTime());

        //Excel file directory path.
        String excelFilePath = "C:/Users/User/Desktop/Assignment1_Java-main/your_file.xlsx";
        FileInputStream excelFile = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(excelFile);

        // Get the sheet based on the current day
        Sheet sheet = workbook.getSheet(currentDay);

        if (sheet != null) {
            for (int rowNumber = 2; rowNumber <= sheet.getLastRowNum(); rowNumber++) {
                Row row = sheet.getRow(rowNumber);
                Cell keywordCell = row.getCell(2);
                String keyword = keywordCell.getStringCellValue();

                driver.get("https://www.google.com");
                WebElement searchBox = driver.findElement(By.name("q"));

                searchBox.clear();
                searchBox.sendKeys(keyword);

                // Wait for auto-suggestions to appear
                wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector(".erkvQe li")));
                List<WebElement> suggestions = driver.findElements(By.cssSelector(".erkvQe li"));

                // Capture the suggested values
                String[] suggestedValues = new String[suggestions.size()];
                for (int i = 0; i < suggestions.size(); i++) {
                    suggestedValues[i] = suggestions.get(i).findElement(By.xpath(".//div")).getText();
                }

                // Get the longest and shortest suggested values
                String longestSuggested = getLongestSuggested(suggestedValues);
                String shortestSuggested = getShortestSuggested(suggestedValues);

                // Update results in the existing sheet
                Cell longestCell = row.createCell(3);  //In the xl sheet Longest Option is starting in the third column.
                longestCell.setCellValue(longestSuggested);
                Cell shortestCell = row.createCell(4);  //In the xl sheet Shortest Option is starting in the fourth column.
                shortestCell.setCellValue(shortestSuggested);
                // Introduce a delay of 5 seconds before the next search
                try {
                    Thread.sleep(3000); // 5000 milliseconds = 5 seconds
                } catch (InterruptedException e) {
                    throw new RuntimeException(e);
                }

            }

            // Save the updated results to the Excel file
            FileOutputStream outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        }


        try {
            driver.quit();
        }  //Exiting from the driver after successfully searched.
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    //The function for calculating and sending the Longest path.
    private static String getLongestSuggested(String[] suggestedValues) {
        String longest = suggestedValues[0];
        for (String suggestion : suggestedValues) {
            if (suggestion.length() > longest.length()) {
                longest = suggestion;
            }
        }
        return longest;
    }

    //The function for calculating and sending the Shortest path.
    private static String getShortestSuggested(String[] suggestedValues) {
        String shortest = suggestedValues[0];
        for (String suggestion : suggestedValues) {
            if (suggestion.length() < shortest.length()) {
                shortest = suggestion;
            }
        }
        return shortest;
    }
}