package org.example;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import java.time.Duration;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.Row;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        System.setProperty("webdriver.chrome.driver", "C:\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        // Create Workbook and Sheet
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Companies Data");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Company Name");
        headerRow.createCell(1).setCellValue("Secteur d'Activité");
        headerRow.createCell(2).setCellValue("Forme Juridique");
        headerRow.createCell(3).setCellValue("Capital");

        int rowNum = 1; // Start from row 1 (after header)

        try {
            // First navigate to get total count
            driver.get("https://www.charika.ma");
            performSearch(driver, wait);

            String companyXPath = "//h5[@class='strong text-lowercase truncate']/a[@class='goto-fiche']";
            
            // Wait for initial results to load
            wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(companyXPath)));
            
            // Get total count of companies
            List<WebElement> companyLinks = driver.findElements(By.xpath(companyXPath));
            int totalCompanies = companyLinks.size();
            System.out.println("Found " + totalCompanies + " companies to process");
            
            for (int i = 0; i < totalCompanies; i++) {
                try {
                    System.out.println("\nProcessing company " + (i + 1));
                    
                    // Start fresh from homepage for each company
                    driver.get("https://www.charika.ma");
                    performSearch(driver, wait);
                    
                    // Wait for elements and click the i-th company
                    wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath(companyXPath)));
                    WebElement companyLink = wait.until(ExpectedConditions.elementToBeClickable(
                        driver.findElements(By.xpath(companyXPath)).get(i)
                    ));
                    
                    String companyName = companyLink.getText();
                    System.out.println("Clicking on company: " + companyName);
                    
                    // Click the company link
                    companyLink.click();
                    
                    // Wait a moment for the page to load
                    Thread.sleep(2000);
                    
                    // Extract company information with null handling
                    String secteurActivite = "N/A";
                    String formeJuridique = "N/A";
                    String capital = "N/A";

                    try {
                        secteurActivite = driver.findElement(
                            By.xpath("//*[@id=\"fiche\"]/div[1]/div[1]/div/div[2]/div[1]/div[2]/span/h2")).getText();
                    } catch (Exception e) {
                        System.out.println("Secteur d'Activité not found for: " + companyName);
                    }
                    
                    try {
                        formeJuridique = driver.findElement(
                            By.xpath("//*[@id=\"fiche\"]/div[1]/div[1]/div/div[2]/div[4]/div/div[1]/table/tbody/tr[3]/td[2]")).getText();
                    } catch (Exception e) {
                        System.out.println("Forme Juridique not found for: " + companyName);
                    }
                    
                    try {
                        capital = driver.findElement(
                            By.xpath("//*[@id=\"fiche\"]/div[1]/div[1]/div/div[2]/div[4]/div/div[1]/table/tbody/tr[4]/td[2]")).getText();
                    } catch (Exception e) {
                        System.out.println("Capital not found for: " + companyName);
                    }

                    // Create a new row in Excel and populate it
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(companyName);
                    row.createCell(1).setCellValue(secteurActivite);
                    row.createCell(2).setCellValue(formeJuridique);
                    row.createCell(3).setCellValue(capital);

                    // Print for console tracking
                    System.out.println("Added to Excel: " + companyName);
                    System.out.println("----------------------------------------");

                } catch (Exception e) {
                    System.out.println("Error processing company " + (i + 1) + ": " + e.getMessage());
                    e.printStackTrace();
                    continue;
                }
            }

            // Write the workbook to a file
            try (FileOutputStream outputStream = new FileOutputStream("MarrakechCompanies.xlsx")) {
                workbook.write(outputStream);
            }
            System.out.println("Excel file has been created successfully!");

        } catch (Exception e) {
            System.out.println("Error occurred: " + e.getMessage());
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            driver.quit();
        }
    }

    private static void performSearch(WebDriver driver, WebDriverWait wait) {
        // Click region selector
        WebElement region = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("//*[@id=\"national\"]/form/div/div[2]/div/div/div/button/div/div/div")));
        region.click();
        
        // Enter city
        WebElement ville = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("//*[@id=\"national\"]/form/div/div[2]/div/div/div/div/div[1]/input")));
        ville.sendKeys("Marrakech");
        ville.submit();
    }
}