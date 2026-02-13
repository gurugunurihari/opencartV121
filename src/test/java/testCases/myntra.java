package testCases;

import java.io.FileOutputStream;
import java.time.Duration;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;

public class myntra {

    public static void main(String[] args) throws Exception {

        WebDriver driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
        driver.manage().window().maximize();

        
        driver.get("https://www.myntra.com");

        
        driver.findElement(By.xpath("//input[@placeholder='Search for products, brands and more']"))
              .sendKeys("shirts", Keys.ENTER);

     
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Myntra Products");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Product Name");
        header.createCell(1).setCellValue("Price");

        int rowNum = 1;

       
        int highestPrice = 0;
        WebElement highestPriceProduct = null;

        
        int productCount = driver.findElements(
                By.xpath("//li[@class='product-base']"))
                .size();

        for (int i = 1; i <= productCount; i++) {
            try {
                WebElement product = driver.findElement(
                        By.xpath("(//li[@class='product-base'])[" + i + "]"));

                String name = product.findElement(
                        By.xpath(".//h4[@class='product-product']"))
                        .getText();

                String priceText = product.findElement(
                        By.xpath(".//span[@class='product-discountedPrice'] | .//span[@class='product-price']"))
                        .getText()
                        .replace("Rs. ", "")
                        .replace(",", "");

                int price = Integer.parseInt(priceText);

             
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(name);
                row.createCell(1).setCellValue(price);

                if (price > highestPrice) {
                    highestPrice = price;
                    highestPriceProduct = product;
                }

            } catch (Exception e) {
               
        }

      
        FileOutputStream fos =
                new FileOutputStream("MyntraProducts.xlsx");
        workbook.write(fos);
        workbook.close();
        fos.close();

     
        highestPriceProduct.click();

        for (String window : driver.getWindowHandles()) {
            driver.switchTo().window(window);
        }

       
        driver.findElement(
                By.xpath("(//button[contains(@class,'size-buttons-size-button')])[1]"))
                .click();

    
        driver.findElement(
                By.xpath("//div[text()='ADD TO BAG']"))
                .click();

        driver.findElement(
                By.xpath("//a[@href='/checkout/cart']"))
                .click();

        System.out.println(
                "Highest price product added to bag: Rs." + highestPrice);
    }
    }
}

