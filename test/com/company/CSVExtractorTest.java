package com.company;

import org.junit.jupiter.api.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

class CSVExtractorTest {

    static List<String> outFilenames;

    @BeforeAll
    static void SetUp() {
        outFilenames = new ArrayList<>();
    }

    @AfterAll
    static void TearDown() throws IOException {

        for (String s : outFilenames) {
            if (Files.exists(Paths.get(s))) {
                Files.delete(Paths.get(s));
            }
        }
        Files.delete(Paths.get("test/csv/"));
        Files.delete(Paths.get("test/TestDirectory/csv/"));
    }

    @Test
    void TestExistingFile() throws IOException {
        CSVExtractor.main(new String[] {"test/TestExistingFile.xlsx"});
        AssertFileExists("test/csv/Data.csv");
    }

    @Test
    void TestMissingFile() {
        try {
            CSVExtractor.main(new String[]{"test/MissingFile.xlsx"});
        } catch (Exception e) {
            Assertions.fail("No exception should be thrown if the file is missing.");
        }
    }

    @Test
    void TestNoCsv() throws IOException {
        try {
            CSVExtractor.main(new String[] {"test/TestNoCSVs.xlsx"});
        } catch (Exception e) {
            Assertions.fail("No exception should be thrown if no CSVs have been specified in the file.");
        }
    }

    @Test
    void TestMultipleCsv() throws IOException {
        CSVExtractor.main(new String[] {"test/TestMultipleCSVs.xlsx"});
        for (int i = 1; i <= 3; i++) {
            AssertFileExists(String.format("test/csv/Data%d.csv", i));
        }
    }

    @Test
    void TestMultipleSheets() throws IOException {
        CSVExtractor.main(new String[] {"test/TestMultipleXLSSheets.xlsx"});
        for (int i = 1; i <= 9; i++) {
            AssertFileExists(String.format("test/csv/Data%d.csv", i));
        }
    }

    @Test
    void TestMultipleXls() throws IOException {
        CSVExtractor.main(new String[] {"test/TestMultipleXLSs1.xlsx", "test/TestMultipleXLSs2.xlsx", "test/TestMultipleXLSs3.xlsx"});
        for (int i = 1; i <= 3; i++) {
            AssertFileExists(String.format("test/csv/Data%d.csv", i));
        }
    }

    @Test
    void TestDeeperDirectory() throws IOException {
        CSVExtractor.main(new String[] {"test/TestDirectory/TestDeeperDirectory.xlsx"});
        AssertFileExists("test/TestDirectory/csv/Data.csv");
    }

    @Test
    void TestIncompleteCsv() throws IOException {
        CSVExtractor.main(new String[] {"test/TestIncompleteCSV.xlsx"});
        Assertions.assertFalse(Files.exists(Paths.get("test/csv/IncompleteData.csv")));
    }

    @Test
    void TestNoCSVThicknessSpecified() throws IOException {
        CSVExtractor.main(new String[] {"test/TestNoCSVThicknessSpecified.xlsx"});
        Assertions.assertFalse(Files.exists(Paths.get("test/csv/Data.csv")));
    }

    @Test
    void TestVerifyContents() throws IOException {
        CSVExtractor.main(new String[] {"test/TestVerifyContents.xlsx"});
        AssertFileExists("test/csv/TestVerifyContents.csv");
        Assertions.assertArrayEquals(Files.readAllLines(Paths.get("test/csv/TestVerifyContents.csv")).toArray(), new String[] { "1,2,3", "4,5,6", "7,8,9" });
    }

    @Test
    void TestUnicodeCharacters() throws IOException {
        CSVExtractor.main(new String[] {"test/TestUnicodeCharacters.xlsx"});
        Assertions.assertEquals(Files.readAllLines(Paths.get("test/csv/MenuExample.csv")).get(0), "skillet beef,tiěbǎn niúròu,tyeah-ban nyoh-roh,铁板牛肉");
        AssertFileExists("test/csv/MenuExample.csv");
    }

    @Test
    void TestFormula() throws IOException {
        CSVExtractor.main(new String[] {"test/TestFormula.xlsx"});
        Assertions.assertEquals(Files.readAllLines(Paths.get("test/csv/Formula.csv")).get(28), "329.11154727235277,32984.056261608435" );
        AssertFileExists("test/csv/Formula.csv");
    }

    private static void AssertFileExists(String fileName) {
        outFilenames.add(fileName);
        Assertions.assertTrue(Files.exists(Paths.get(fileName)));
    }
}