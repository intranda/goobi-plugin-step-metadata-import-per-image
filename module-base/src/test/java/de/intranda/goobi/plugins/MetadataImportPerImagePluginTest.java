package de.intranda.goobi.plugins;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.mockConstruction;
import static org.mockito.Mockito.mockStatic;
import static org.mockito.Mockito.when;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.goobi.beans.Process;
import org.goobi.beans.Project;
import org.goobi.beans.Ruleset;
import org.goobi.beans.Step;
import org.goobi.beans.User;
import org.goobi.production.enums.PluginReturnValue;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.io.TempDir;
import org.mockito.MockedConstruction;
import org.mockito.MockedStatic;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

import de.sub.goobi.config.ConfigurationHelper;
import de.sub.goobi.helper.Helper;
import de.sub.goobi.helper.StorageProvider;
import de.sub.goobi.helper.StorageProviderInterface;
import de.sub.goobi.helper.VariableReplacer;
import de.sub.goobi.helper.enums.StepStatus;
import de.sub.goobi.metadaten.MetadatenHelper;
import de.sub.goobi.metadaten.MetadatenImagesHelper;
import de.sub.goobi.persistence.managers.MetadataManager;
import de.sub.goobi.persistence.managers.ProcessManager;
import ugh.dl.DocStruct;
import ugh.dl.Fileformat;
import ugh.dl.Metadata;
import ugh.dl.MetadataType;
import ugh.dl.Prefs;
import ugh.dl.Reference;
import ugh.fileformats.mets.MetsMods;

@ExtendWith(MockitoExtension.class)
// LENIENT: tests such as testConstructor/testInit and the validate* tests exercise only part of the plugin,
// leaving some of the shared setUp() stubs unused in those individual tests.
@MockitoSettings(strictness = Strictness.LENIENT)
public class MetadataImportPerImagePluginTest {

    private static String resourcesFolder;

    @TempDir
    File tempDir;

    private File processDirectory;
    private File metadataDirectory;
    private Process process;
    private Step step;
    private Prefs prefs;
    private Path metaTarget;

    private MockedStatic<ConfigurationHelper> mockedConfigHelper;
    private MockedStatic<VariableReplacer> mockedVariableReplacer;
    private MockedStatic<MetadatenHelper> mockedMetadatenHelper;
    private MockedStatic<MetadataManager> mockedMetadataManager;
    private MockedStatic<ProcessManager> mockedProcessManager;
    private MockedStatic<Helper> mockedHelper;
    private MockedStatic<StorageProvider> mockedStorageProvider;
    private MockedConstruction<MetadatenImagesHelper> mockedImagesHelper;

    @BeforeAll
    public static void setUpClass() throws Exception {
        resourcesFolder = "src/test/resources/";

        if (!Files.exists(Paths.get(resourcesFolder))) {
            resourcesFolder = "target/test-classes/";
        }

        String log4jFile = resourcesFolder + "log4j2.xml";
        System.setProperty("log4j.configurationFile", log4jFile);
    }

    @Test
    public void testConstructor() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        assertNotNull(plugin);
    }

    @Test
    public void testInit() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        assertEquals(step.getTitel(), plugin.getStep().getTitel());
    }

    @Test
    public void testInitReadsPaginationLabelMetadata() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        assertEquals("logicalPageNumber", plugin.paginationLabelMetadata);
    }

    @Test
    public void testParseExcel() throws Exception {
        // Create a test Excel file with known content
        File excelFile = new File(tempDir, "test.xlsx");
        createTestExcel(excelFile, createTestRows());

        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.columnLabel = "Label";

        MetadataImportPerImageStepPlugin.ParseResult result = plugin.parseExcel(excelFile.getAbsolutePath());

        assertEquals(10, result.rows().size());
        assertEquals("uri1", result.rows().get(0).get("URI"));
        assertEquals("folder1", result.rows().get(0).get("Structure"));
        assertEquals("p1", result.rows().get(0).get("Label"));
        assertEquals("caption1", result.rows().get(0).get("Caption"));
        // Row with empty structure
        assertEquals("", result.rows().get(8).get("Structure"));
    }

    @Test
    public void testParseExcelWithFormulaErrorCell() throws Exception {
        File excelFile = new File(tempDir, "error.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("URI");
            header.createCell(1).setCellValue("Structure");
            header.createCell(2).setCellValue("Label");
            header.createCell(3).setCellValue("Caption");

            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("uri1");
            // Cell with an explicit error value
            dataRow.createCell(1).setCellErrorValue(FormulaError.DIV0.getCode());
            dataRow.createCell(2).setCellValue("p1");
            dataRow.createCell(3).setCellValue("caption1");

            try (java.io.FileOutputStream fos = new java.io.FileOutputStream(excelFile)) {
                workbook.write(fos);
            }
        }

        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.columnLabel = "Label";

        MetadataImportPerImageStepPlugin.ParseResult result = plugin.parseExcel(excelFile.getAbsolutePath());
        assertEquals(1, result.rows().size());
        // Error cell should produce an empty string
        assertEquals("", result.rows().get(0).get("Structure"));
        assertEquals("uri1", result.rows().get(0).get("URI"));
    }

    @Test
    public void testBuildStructureCreatesHierarchy() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.readConfiguration(step);
        plugin.columnLabel = "Label";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        List<Map<String, String>> rows = createTestRows();

        plugin.buildStructure(ff, prefs, rows);

        // The logical root is MultiVolumeWork (anchor); content is under Volume
        DocStruct anchor = ff.getDigitalDocument().getLogicalDocStruct();
        DocStruct volume = anchor.getAllChildren().get(0);

        // Expect 2 Chapter elements (uri1, uri2)
        assertNotNull(volume.getAllChildren());
        assertEquals(2, volume.getAllChildren().size());

        DocStruct chapterUri1 = volume.getAllChildren().get(0);
        assertEquals("Chapter", chapterUri1.getType().getName());
        assertEquals("uri1", chapterUri1.getAllMetadata().get(0).getValue());

        // uri1 should have 2 sub-Chapters (folder1, folder2)
        assertEquals(2, chapterUri1.getAllChildren().size());
        DocStruct folderChapter1 = chapterUri1.getAllChildren().get(0);
        assertEquals("folder1", folderChapter1.getAllMetadata().get(0).getValue());

        // folder1 should have 2 Figures (caption1, caption2)
        assertEquals(2, folderChapter1.getAllChildren().size());

        DocStruct chapterUri2 = volume.getAllChildren().get(1);
        // uri2 should have 2 sub-Chapters (folder3, Unnamed folder)
        assertEquals(2, chapterUri2.getAllChildren().size());
        DocStruct unnamedFolder = chapterUri2.getAllChildren().get(1);
        assertEquals("Unnamed folder", unnamedFolder.getAllMetadata().get(0).getValue());
    }

    @Test
    public void testBuildStructureSetsPaginationLabel() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.readConfiguration(step);
        plugin.columnLabel = "Label";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        List<Map<String, String>> rows = createTestRows();
        plugin.buildStructure(ff, prefs, rows);

        List<DocStruct> pages = ff.getDigitalDocument().getPhysicalDocStruct().getAllChildren();
        MetadataType logicalPageNumber = prefs.getMetadataTypeByName("logicalPageNumber");
        assertEquals("p1", pages.get(0).getAllMetadataByType(logicalPageNumber).get(0).getValue());
        assertEquals("p5", pages.get(4).getAllMetadataByType(logicalPageNumber).get(0).getValue());
        assertEquals("p10", pages.get(9).getAllMetadataByType(logicalPageNumber).get(0).getValue());
    }

    @Test
    public void testBuildStructureClearsExistingPaginationLabelWhenEmpty() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.readConfiguration(step);
        plugin.columnLabel = "Label";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        // Verify page 3 (index 2) has an existing logicalPageNumber ("10" from meta.xml ORDERLABEL)
        List<DocStruct> pagesBefore = ff.getDigitalDocument().getPhysicalDocStruct().getAllChildren();
        MetadataType logicalPageNumber = prefs.getMetadataTypeByName("logicalPageNumber");
        String originalLabel = pagesBefore.get(2).getAllMetadataByType(logicalPageNumber).get(0).getValue();
        assertEquals("10", originalLabel);

        // Create test rows where row index 2 has an empty label
        List<Map<String, String>> rows = createTestRows();
        rows.set(2, row("uri1", "folder1", "", "caption2"));

        plugin.buildStructure(ff, prefs, rows);

        // The page with the empty label should have its logicalPageNumber set to ""
        List<DocStruct> pagesAfter = ff.getDigitalDocument().getPhysicalDocStruct().getAllChildren();
        String clearedLabel = pagesAfter.get(2).getAllMetadataByType(logicalPageNumber).get(0).getValue();
        assertEquals("", clearedLabel);

        // Other pages should still have their labels set from Excel
        assertEquals("p1", pagesAfter.get(0).getAllMetadataByType(logicalPageNumber).get(0).getValue());
        assertEquals("p4", pagesAfter.get(3).getAllMetadataByType(logicalPageNumber).get(0).getValue());
    }

    @Test
    public void testBuildStructureLinksPages() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.readConfiguration(step);
        plugin.columnLabel = "Label";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        List<Map<String, String>> rows = createTestRows();
        plugin.buildStructure(ff, prefs, rows);

        DocStruct anchor = ff.getDigitalDocument().getLogicalDocStruct();
        DocStruct volume = anchor.getAllChildren().get(0);

        // Volume should be linked to all 10 pages
        assertEquals(10, volume.getAllToReferences().size());

        // First Chapter (uri1) should be linked to first 5 pages
        DocStruct chapterUri1 = volume.getAllChildren().get(0);
        assertEquals(5, chapterUri1.getAllToReferences().size());
    }

    @Test
    public void testBuildStructureReplacesExistingChildren() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.readConfiguration(step);
        plugin.columnLabel = "Label";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        // Original meta.xml has 3 Figure children under Volume
        DocStruct volume = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        assertEquals(3, volume.getAllChildren().size());

        List<Map<String, String>> rows = createTestRows();
        plugin.buildStructure(ff, prefs, rows);

        // After rebuild, Volume should have 2 Chapter children (uri1, uri2)
        assertEquals(2, volume.getAllChildren().size());
        assertEquals("Chapter", volume.getAllChildren().get(0).getType().getName());
    }

    @Test
    public void testBuildStructureRemovesStaleReferences() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.readConfiguration(step);
        plugin.columnLabel = "Label";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        // First call: build structure with the standard test rows
        List<Map<String, String>> rows1 = createTestRows();
        plugin.buildStructure(ff, prefs, rows1);

        // Collect the logical structures created by the first call
        DocStruct volume = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        List<DocStruct> firstCallChildren = new ArrayList<>(volume.getAllChildren());

        // Second call: rebuild with different data (swap uri values)
        List<Map<String, String>> rows2 = new ArrayList<>();
        for (Map<String, String> r : rows1) {
            Map<String, String> copy = new HashMap<>(r);
            copy.put("URI", "different_" + copy.get("URI"));
            rows2.add(copy);
        }
        plugin.buildStructure(ff, prefs, rows2);

        // Verify: physical pages should NOT have back-references to structures from the first call
        List<DocStruct> pages = ff.getDigitalDocument().getPhysicalDocStruct().getAllChildren();
        for (DocStruct page : pages) {
            if (page.getAllFromReferences() != null) {
                for (Reference ref : page.getAllFromReferences()) {
                    for (DocStruct oldChild : firstCallChildren) {
                        if (ref.getSource() == oldChild) {
                            throw new AssertionError("Physical page still references a stale logical structure");
                        }
                    }
                }
            }
        }

        // Also verify the new structure was built correctly
        assertEquals(2, volume.getAllChildren().size());
        assertEquals("different_uri1", volume.getAllChildren().get(0).getAllMetadata().get(0).getValue());
    }

    @Test
    public void testBuildStructurePreservesExistingMetadata() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.readConfiguration(step);
        plugin.columnLabel = "Label";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        List<Map<String, String>> rows = createTestRows();

        // First build: creates the hierarchy from scratch
        plugin.buildStructure(ff, prefs, rows);

        DocStruct volume = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        DocStruct chapterUri1 = volume.getAllChildren().get(0);
        assertEquals("uri1", chapterUri1.getAllMetadata().get(0).getValue());

        // Simulate external metadata added after first import (e.g. by a user or another plugin)
        MetadataType shortTitleType = prefs.getMetadataTypeByName("TitleDocMainShort");
        Metadata extraMd = new Metadata(shortTitleType);
        extraMd.setValue("important value that must survive re-import");
        chapterUri1.addMetadata(extraMd);

        // Second build: same Excel data — should reuse existing elements and preserve metadata
        plugin.buildStructure(ff, prefs, rows);

        DocStruct volumeAfter = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        assertEquals(2, volumeAfter.getAllChildren().size());

        DocStruct chapterUri1After = volumeAfter.getAllChildren().get(0);
        assertEquals("Chapter", chapterUri1After.getType().getName());
        assertEquals("uri1", chapterUri1After.getAllMetadataByType(
                prefs.getMetadataTypeByName("TitleDocMain")).get(0).getValue());

        // The externally added metadata must still be present
        List<? extends Metadata> shortTitles = chapterUri1After.getAllMetadataByType(shortTitleType);
        assertFalse(shortTitles.isEmpty(), "Existing metadata should be preserved after re-import");
        assertEquals("important value that must survive re-import", shortTitles.get(0).getValue());
    }

    @Test
    public void testBuildStructureWithoutPhysicalPages() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.columnLabel = "Label";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        // Remove all physical page children
        DocStruct physRoot = ff.getDigitalDocument().getPhysicalDocStruct();
        if (physRoot.getAllChildren() != null) {
            for (DocStruct child : new ArrayList<>(physRoot.getAllChildren())) {
                physRoot.removeChild(child);
            }
        }

        List<Map<String, String>> rows = createTestRows();
        assertThrows(IllegalStateException.class, () -> plugin.buildStructure(ff, prefs, rows));
    }

    @Test
    public void testBuildStructureThrowsForAnchorWithoutChildren() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.columnLabel = "Label";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        // Remove all children from the anchor
        DocStruct anchor = ff.getDigitalDocument().getLogicalDocStruct();
        if (anchor.getAllChildren() != null) {
            for (DocStruct child : new ArrayList<>(anchor.getAllChildren())) {
                anchor.removeChild(child);
            }
        }

        List<Map<String, String>> rows = createTestRows();
        assertThrows(IllegalStateException.class, () -> plugin.buildStructure(ff, prefs, rows));
    }

    @Test
    public void testBuildStructureHandlesNullLogicalPageNumberType() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.readConfiguration(step);
        plugin.columnLabel = "Label";
        plugin.paginationLabelMetadata = "nonExistentMetadataType";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        List<Map<String, String>> rows = createTestRows();
        plugin.buildStructure(ff, prefs, rows);

        // Hierarchy structure should still be built correctly
        DocStruct anchor = ff.getDigitalDocument().getLogicalDocStruct();
        DocStruct volume = anchor.getAllChildren().get(0);
        assertNotNull(volume.getAllChildren());
        assertEquals(2, volume.getAllChildren().size());

        // Pagination labels should NOT have been updated to Excel values
        List<DocStruct> pages = ff.getDigitalDocument().getPhysicalDocStruct().getAllChildren();
        MetadataType logicalPageNumber = prefs.getMetadataTypeByName("logicalPageNumber");
        String firstPageLabel = pages.get(0).getAllMetadataByType(logicalPageNumber).get(0).getValue();
        assertNotNull(firstPageLabel);
        // Should still have original value, not "p1" from Excel
        assertEquals("8", firstPageLabel);
    }

    @Test
    public void testValidateExcelDataCollectsMultipleErrors() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.hierarchyLevels = new ArrayList<>();

        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "Chapter", "URI", "TitleDocMain", "", "", "exact", "metadataMatchesExcel"));
        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "Chapter", "MissingColumn", "TitleDocMain", "", "", "exact", "metadataMatchesExcel"));

        // Rows that don't have "MissingColumn" and mismatched counts
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row = new HashMap<>();
        row.put("URI", "uri1");
        row.put("Label", "p1");
        rows.add(row);

        // imageCount=5 != rows.size()=1, physPageCount=3 != rows.size()=1, missing column "MissingColumn"
        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateExcelData(rows, 5, 3, 1);

        assertTrue(result.getErrors().size() >= 2, "Expected multiple errors");
    }

    @Test
    public void testValidateConfigMissingStructType() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.hierarchyLevels = new ArrayList<>();

        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "NonExistentType", "URI", "", "", "", "exact", "metadataMatchesExcel"));

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateConfig(prefs);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("NonExistentType")),
                "Error should mention the type name");
    }

    @Test
    public void testValidateConfigMissingMetadataField() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.hierarchyLevels = new ArrayList<>();

        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "Chapter", "URI", "NonExistentField", "", "", "exact", "metadataMatchesExcel"));

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateConfig(prefs);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("NonExistentField")),
                "Error should mention the field name");
    }

    @Test
    public void testValidateConfigMissingMatchMetadata() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.hierarchyLevels = new ArrayList<>();

        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "Chapter", "URI", "TitleDocMain", "", "NonExistentMatchField", "exact", "metadataMatchesExcel"));

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateConfig(prefs);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("NonExistentMatchField")),
                "Error should mention the match metadata field");
    }

    @Test
    public void testValidateConfigInvalidMatchMode() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.hierarchyLevels = new ArrayList<>();

        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "Chapter", "URI", "TitleDocMain", "", "", "invalidMode", "metadataMatchesExcel"));

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateConfig(prefs);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("invalidMode")),
                "Error should mention invalid matchMode");
    }

    @Test
    public void testValidateConfigInvalidMatchDirection() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.hierarchyLevels = new ArrayList<>();

        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "Chapter", "URI", "TitleDocMain", "", "", "exact", "invalidDirection"));

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateConfig(prefs);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("invalidDirection")),
                "Error should mention invalid matchDirection");
    }

    @Test
    public void testValidateLabelColumnMissing() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.columnLabel = "MissingLabelColumn";
        plugin.hierarchyLevels = new ArrayList<>();

        // Row without the expected label column
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row = new HashMap<>();
        row.put("URI", "uri1");
        rows.add(row);

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateExcelData(rows, 1, 1, 1);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("MissingLabelColumn")),
                "Error should mention the missing label column");
    }

    @Test
    public void testParseDuplicateColumnHeaders() throws Exception {
        File excelFile = new File(tempDir, "duplicate_headers.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("URI");
            header.createCell(1).setCellValue("Structure");
            header.createCell(2).setCellValue("URI"); // duplicate
            header.createCell(3).setCellValue("Label");

            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("uri1");
            dataRow.createCell(1).setCellValue("folder1");
            dataRow.createCell(2).setCellValue("uri_dup");
            dataRow.createCell(3).setCellValue("p1");

            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                workbook.write(fos);
            }
        }

        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();

        IOException e = assertThrows(IOException.class, () -> plugin.parseExcel(excelFile.getAbsolutePath()));
        assertTrue(e.getMessage().contains("Duplicate column header: URI"),
                "Exception message should mention the duplicate header");
    }

    @Test
    public void testValidateEmptyGroupByWithoutFallback() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.columnLabel = "Label";
        plugin.hierarchyLevels = new ArrayList<>();

        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "Chapter", "URI", "TitleDocMain", "", "", "exact", "metadataMatchesExcel")); // no fallback

        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row1 = new HashMap<>();
        row1.put("URI", "uri1");
        row1.put("Label", "p1");
        rows.add(row1);
        Map<String, String> row2 = new HashMap<>();
        row2.put("URI", ""); // empty value, no fallback
        row2.put("Label", "p2");
        rows.add(row2);

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateExcelData(rows, 2, 2, 1);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("Row 3") && e.contains("URI")),
                "Error should mention the row and column");
    }

    @Test
    public void testValidateEmptyExcel() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.columnLabel = "Label";
        plugin.hierarchyLevels = new ArrayList<>();

        List<Map<String, String>> rows = new ArrayList<>(); // empty

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateExcelData(rows, 0, 0, 1);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("no data rows")),
                "Error should mention no data rows");
    }

    @Test
    public void testValidateXmlInvalidChars() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.columnLabel = "Label";
        plugin.hierarchyLevels = new ArrayList<>();

        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row = new HashMap<>();
        row.put("Label", "p1");
        row.put("URI", "urivalue"); // contains control char 0x03
        rows.add(row);

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateExcelData(rows, 1, 1, 1);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("invalid XML control characters")),
                "Error should mention XML control characters");
    }

    @Test
    public void testValidateAllGroupColumnsEmpty() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.columnLabel = "Label";
        plugin.hierarchyLevels = new ArrayList<>();

        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "Chapter", "URI", "TitleDocMain", "Fallback1", "", "exact", "metadataMatchesExcel"));
        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "Chapter", "Structure", "TitleDocMain", "Fallback2", "", "exact", "metadataMatchesExcel"));

        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row = new HashMap<>();
        row.put("URI", "");
        row.put("Structure", "");
        row.put("Label", "p1");
        rows.add(row);

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateExcelData(rows, 1, 1, 1);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("all grouping columns are empty")),
                "Error should mention all grouping columns empty");
    }

    @Test
    public void testValidateNonContiguousDuplicateKeysWarning() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.columnLabel = "Label";
        plugin.hierarchyLevels = new ArrayList<>();

        plugin.hierarchyLevels.add(new MetadataImportPerImageStepPlugin.HierarchyLevel(
                "Chapter", "URI", "TitleDocMain", "", "", "exact", "metadataMatchesExcel"));

        // Pattern: A, B, A — non-contiguous
        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row1 = new HashMap<>();
        row1.put("URI", "A");
        row1.put("Label", "p1");
        rows.add(row1);
        Map<String, String> row2 = new HashMap<>();
        row2.put("URI", "B");
        row2.put("Label", "p2");
        rows.add(row2);
        Map<String, String> row3 = new HashMap<>();
        row3.put("URI", "A");
        row3.put("Label", "p3");
        rows.add(row3);

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateExcelData(rows, 3, 3, 1);

        assertFalse(result.hasErrors(), "Should not have errors");
        assertFalse(result.getWarnings().isEmpty(), "Should have warnings");
        assertTrue(result.getWarnings().stream().anyMatch(w -> w.contains("Non-contiguous") && w.contains("A")),
                "Warning should mention non-contiguous values");
    }

    @Test
    public void testValidateMultipleSheetsWarning() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.columnLabel = "Label";
        plugin.hierarchyLevels = new ArrayList<>();

        List<Map<String, String>> rows = new ArrayList<>();
        Map<String, String> row = new HashMap<>();
        row.put("Label", "p1");
        rows.add(row);

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateExcelData(rows, 1, 1, 3);

        assertFalse(result.hasErrors(), "Should not have errors");
        assertFalse(result.getWarnings().isEmpty(), "Should have warnings");
        assertTrue(result.getWarnings().stream().anyMatch(w -> w.contains("3 sheets")),
                "Warning should mention multiple sheets");
    }

    @Test
    public void testValidateConfigNoHierarchyLevels() {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.hierarchyLevels = new ArrayList<>(); // empty

        MetadataImportPerImageStepPlugin.ValidationResult result = plugin.validateConfig(prefs);

        assertTrue(result.hasErrors(), "Expected errors");
        assertTrue(result.getErrors().stream().anyMatch(e -> e.contains("No hierarchy levels configured")),
                "Error should mention no hierarchy levels");
    }

    @Test
    public void testBuildStructurePositionalMatchingFallback() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.columnLabel = "Label";

        // Override levels to have no matchMetadata — triggers positional matching
        plugin.hierarchyLevels = List.of(
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Chapter", "URI", "TitleDocMain", "", "", "exact", "metadataMatchesExcel"),
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Chapter", "Structure", "TitleDocMain", "Unnamed folder", "", "exact", "metadataMatchesExcel"),
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Figure", "Caption", "TitleDocMain", "", "", "exact", "metadataMatchesExcel"));

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        List<Map<String, String>> rows = createTestRows();

        // First build
        plugin.buildStructure(ff, prefs, rows);
        DocStruct volume = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        DocStruct chapterUri1 = volume.getAllChildren().get(0);

        // Add marker metadata to verify element reuse
        MetadataType shortTitleType = prefs.getMetadataTypeByName("TitleDocMainShort");
        Metadata marker = new Metadata(shortTitleType);
        marker.setValue("positional-marker");
        chapterUri1.addMetadata(marker);

        // Second build: positional matching should reuse the element
        plugin.buildStructure(ff, prefs, rows);
        DocStruct volumeAfter = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        DocStruct chapterUri1After = volumeAfter.getAllChildren().get(0);
        List<? extends Metadata> markers = chapterUri1After.getAllMetadataByType(shortTitleType);
        assertFalse(markers.isEmpty(), "Marker metadata should survive positional reuse");
        assertEquals("positional-marker", markers.get(0).getValue());

        // Warnings should be populated
        assertFalse(plugin.buildWarnings.isEmpty(), "buildWarnings should contain positional matching warnings");
    }

    @Test
    public void testBuildStructureMatchMetadataDifferentFromWriteField() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.columnLabel = "Label";

        // matchMetadata="_ucc_id" but metadataField="TitleDocMain" (write field)
        plugin.hierarchyLevels = List.of(
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Chapter", "URI", "TitleDocMain", "", "_ucc_id", "exact", "metadataMatchesExcel"),
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Chapter", "Structure", "TitleDocMain", "Unnamed folder",
                        "TitleDocMain", "exact", "metadataMatchesExcel"),
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Figure", "Caption", "TitleDocMain", "",
                        "TitleDocMain", "exact", "metadataMatchesExcel"));

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        List<Map<String, String>> rows = createTestRows();

        // First build: creates structure with TitleDocMain set to the group key
        plugin.buildStructure(ff, prefs, rows);
        DocStruct volume = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        DocStruct chapterUri1 = volume.getAllChildren().get(0);

        // Manually add _ucc_id metadata (simulating external enrichment)
        MetadataType uccIdType = prefs.getMetadataTypeByName("_ucc_id");
        Metadata uccId = new Metadata(uccIdType);
        uccId.setValue("uri1");
        chapterUri1.addMetadata(uccId);

        // Add marker to verify reuse
        MetadataType shortTitleType = prefs.getMetadataTypeByName("TitleDocMainShort");
        Metadata marker = new Metadata(shortTitleType);
        marker.setValue("match-by-ucc-id");
        chapterUri1.addMetadata(marker);

        // Second build: should match by _ucc_id, not TitleDocMain
        plugin.buildStructure(ff, prefs, rows);
        DocStruct volumeAfter = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        DocStruct chapterUri1After = volumeAfter.getAllChildren().get(0);
        List<? extends Metadata> markers = chapterUri1After.getAllMetadataByType(shortTitleType);
        assertFalse(markers.isEmpty(), "Marker should survive when matched by _ucc_id");
        assertEquals("match-by-ucc-id", markers.get(0).getValue());
    }

    @Test
    public void testBuildStructureMatchModeEndsWith() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.columnLabel = "Label";

        // matchMetadata="_ucc_id", matchMode="endsWith" (metadata ends with Excel value)
        plugin.hierarchyLevels = List.of(
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Chapter", "URI", "TitleDocMain", "", "_ucc_id", "endsWith", "metadataMatchesExcel"),
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Chapter", "Structure", "TitleDocMain", "Unnamed folder",
                        "TitleDocMain", "exact", "metadataMatchesExcel"),
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Figure", "Caption", "TitleDocMain", "",
                        "TitleDocMain", "exact", "metadataMatchesExcel"));

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        List<Map<String, String>> rows = createTestRows();

        // First build
        plugin.buildStructure(ff, prefs, rows);
        DocStruct volume = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        DocStruct chapterUri1 = volume.getAllChildren().get(0);

        // Set _ucc_id with a prefix — simulating metadata that has a longer path
        MetadataType uccIdType = prefs.getMetadataTypeByName("_ucc_id");
        Metadata uccId = new Metadata(uccIdType);
        uccId.setValue("/aspace/repositories/11/uri1");
        chapterUri1.addMetadata(uccId);

        // Add marker
        MetadataType shortTitleType = prefs.getMetadataTypeByName("TitleDocMainShort");
        Metadata marker = new Metadata(shortTitleType);
        marker.setValue("endswith-match");
        chapterUri1.addMetadata(marker);

        // Second build: Excel has "uri1", metadata has "/aspace/repositories/11/uri1"
        // endsWith should match because metadata.endsWith("uri1") is true
        plugin.buildStructure(ff, prefs, rows);
        DocStruct volumeAfter = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        DocStruct chapterUri1After = volumeAfter.getAllChildren().get(0);
        List<? extends Metadata> markers = chapterUri1After.getAllMetadataByType(shortTitleType);
        assertFalse(markers.isEmpty(), "Marker should survive endsWith matching");
        assertEquals("endswith-match", markers.get(0).getValue());
    }

    @Test
    public void testBuildStructureMatchModeEndsWithExcelMatchesMetadata() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.columnLabel = "Label";

        // matchDirection="excelMatchesMetadata": Excel value is tested with endsWith against metadata
        plugin.hierarchyLevels = List.of(
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Chapter", "URI", "TitleDocMain", "", "_ucc_id", "endsWith", "excelMatchesMetadata"),
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Chapter", "Structure", "TitleDocMain", "Unnamed folder",
                        "TitleDocMain", "exact", "metadataMatchesExcel"),
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Figure", "Caption", "TitleDocMain", "",
                        "TitleDocMain", "exact", "metadataMatchesExcel"));

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        // Excel URIs have a prefix; metadata stores just the short ID
        List<Map<String, String>> rows = new ArrayList<>();
        rows.add(row("/aspace/repo/uri1", "folder1", "p1", "caption1"));
        rows.add(row("/aspace/repo/uri1", "folder1", "p2", "caption1"));
        rows.add(row("/aspace/repo/uri1", "folder1", "p3", "caption2"));
        rows.add(row("/aspace/repo/uri1", "folder2", "p4", "caption3"));
        rows.add(row("/aspace/repo/uri1", "folder2", "p5", "caption3"));
        rows.add(row("/aspace/repo/uri2", "folder3", "p6", "caption4"));
        rows.add(row("/aspace/repo/uri2", "folder3", "p7", "caption4"));
        rows.add(row("/aspace/repo/uri2", "folder3", "p8", "caption5"));
        rows.add(row("/aspace/repo/uri2", "", "p9", "caption6"));
        rows.add(row("/aspace/repo/uri2", "", "p10", "caption6"));

        // First build
        plugin.buildStructure(ff, prefs, rows);
        DocStruct volume = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        DocStruct chapterUri1 = volume.getAllChildren().get(0);

        // Set _ucc_id to the short value — metadata stores "uri1"
        MetadataType uccIdType = prefs.getMetadataTypeByName("_ucc_id");
        Metadata uccId = new Metadata(uccIdType);
        uccId.setValue("uri1");
        chapterUri1.addMetadata(uccId);

        // Add marker
        MetadataType shortTitleType = prefs.getMetadataTypeByName("TitleDocMainShort");
        Metadata marker = new Metadata(shortTitleType);
        marker.setValue("reverse-endswith");
        chapterUri1.addMetadata(marker);

        // Second build: Excel "/aspace/repo/uri1".endsWith("uri1") -> true
        plugin.buildStructure(ff, prefs, rows);
        DocStruct volumeAfter = ff.getDigitalDocument().getLogicalDocStruct().getAllChildren().get(0);
        DocStruct chapterUri1After = volumeAfter.getAllChildren().get(0);
        List<? extends Metadata> markers = chapterUri1After.getAllMetadataByType(shortTitleType);
        assertFalse(markers.isEmpty(), "Marker should survive excelMatchesMetadata endsWith matching");
        assertEquals("reverse-endswith", markers.get(0).getValue());
    }

    @Test
    public void testPositionalMatchingWarningOncePerLevel() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.columnLabel = "Label";

        // All levels without matchMetadata
        plugin.hierarchyLevels = List.of(
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Chapter", "URI", "TitleDocMain", "", "", "exact", "metadataMatchesExcel"),
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Chapter", "Structure", "TitleDocMain", "Unnamed folder", "", "exact", "metadataMatchesExcel"),
                new MetadataImportPerImageStepPlugin.HierarchyLevel(
                        "Figure", "Caption", "TitleDocMain", "", "", "exact", "metadataMatchesExcel"));

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        List<Map<String, String>> rows = createTestRows();
        plugin.buildStructure(ff, prefs, rows);

        // Should have exactly 3 warnings (one per level), not one per row
        assertEquals(3, plugin.buildWarnings.size(), "Should have exactly one warning per level");
        assertTrue(plugin.buildWarnings.stream().anyMatch(w -> w.contains("Chapter")));
        assertTrue(plugin.buildWarnings.stream().anyMatch(w -> w.contains("Figure")));
    }

    @Test
    public void testRunEndToEnd() throws Exception {
        File excelFile = new File(tempDir, "import.xlsx");
        createTestExcel(excelFile, createTestRows());

        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
        plugin.excelFilePath = excelFile.getAbsolutePath();
        plugin.columnLabel = "Label";

        PluginReturnValue result = plugin.run();

        assertEquals(PluginReturnValue.FINISH, result);

        // Read back the written metadata and verify structure
        Fileformat written = new MetsMods(prefs);
        written.read(metaTarget.toString());
        DocStruct anchor = written.getDigitalDocument().getLogicalDocStruct();
        DocStruct volume = anchor.getAllChildren().get(0);

        // Should have 2 top-level chapters (uri1, uri2)
        assertNotNull(volume.getAllChildren());
        assertEquals(2, volume.getAllChildren().size());
        assertEquals("Chapter", volume.getAllChildren().get(0).getType().getName());

        // Verify page references: volume linked to all 10 pages
        assertEquals(10, volume.getAllToReferences().size());
    }

    @BeforeEach
    public void setUp() throws Exception {
        metadataDirectory = new File(tempDir, "metadata");
        metadataDirectory.mkdirs();
        processDirectory = new File(metadataDirectory + File.separator + "1");
        processDirectory.mkdirs();
        String metadataDirectoryName = metadataDirectory.getAbsolutePath() + File.separator;

        Path metaSource = Paths.get(resourcesFolder, "meta.xml");
        metaTarget = Paths.get(processDirectory.getAbsolutePath(), "meta.xml");
        Files.copy(metaSource, metaTarget);

        Path anchorSource = Paths.get(resourcesFolder, "meta_anchor.xml");
        Path anchorTarget = Paths.get(processDirectory.getAbsolutePath(), "meta_anchor.xml");
        Files.copy(anchorSource, anchorTarget);

        ConfigurationHelper configurationHelper = mock(ConfigurationHelper.class);
        when(configurationHelper.getMetsEditorLockingTime()).thenReturn(1800000L);
        when(configurationHelper.isAllowWhitespacesInFolder()).thenReturn(false);
        when(configurationHelper.useS3()).thenReturn(false);
        when(configurationHelper.isUseProxy()).thenReturn(false);
        when(configurationHelper.getGoobiContentServerTimeOut()).thenReturn(60000);
        when(configurationHelper.getMetadataFolder()).thenReturn(metadataDirectoryName);
        when(configurationHelper.getRulesetFolder()).thenReturn(resourcesFolder);
        when(configurationHelper.getProcessImagesMainDirectoryName()).thenReturn("00469418X_media");
        when(configurationHelper.isUseMasterDirectory()).thenReturn(true);
        when(configurationHelper.isCreateMasterDirectory()).thenReturn(false);
        when(configurationHelper.getProcessImagesMasterDirectoryName()).thenReturn("00469418X_master");
        when(configurationHelper.getProcessImagesFallbackDirectoryName()).thenReturn("");
        when(configurationHelper.getConfigurationFolder()).thenReturn(resourcesFolder);
        when(configurationHelper.getNumberOfMetaBackups()).thenReturn(0);
        when(configurationHelper.getScriptCreateDirMeta()).thenReturn("");
        when(configurationHelper.getGoobiFolder()).thenReturn(resourcesFolder);
        when(configurationHelper.getScriptsFolder()).thenReturn("");

        mockedConfigHelper = mockStatic(ConfigurationHelper.class);
        mockedConfigHelper.when(ConfigurationHelper::getInstance).thenReturn(configurationHelper);

        mockedVariableReplacer = mockStatic(VariableReplacer.class);
        mockedVariableReplacer.when(() -> VariableReplacer.simpleReplace(anyString(), any()))
                .thenAnswer(invocation -> invocation.getArgument(0));

        prefs = new Prefs();
        prefs.loadPrefs(resourcesFolder + "ruleset.xml");
        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        mockedMetadatenHelper = mockStatic(MetadatenHelper.class);
        mockedMetadatenHelper.when(() -> MetadatenHelper.getMetaFileType(anyString())).thenReturn("mets");
        mockedMetadatenHelper.when(() -> MetadatenHelper.getFileformatByName(anyString(), any())).thenReturn(ff);
        mockedMetadatenHelper.when(() -> MetadatenHelper.getMetadataOfFileformat(any()))
                .thenReturn(Collections.emptyMap());

        // Prevent database and journal calls triggered by process.writeMetadataFile() and error reporting.
        // The mocked static void methods are no-ops by default.
        mockedMetadataManager = mockStatic(MetadataManager.class);
        mockedProcessManager = mockStatic(ProcessManager.class);
        mockedHelper = mockStatic(Helper.class);

        StorageProviderInterface spMock = mock(StorageProviderInterface.class);
        when(spMock.isFileExists(any())).thenReturn(true);
        when(spMock.listFiles(anyString())).thenAnswer(invocation -> {
            String dirPath = invocation.getArgument(0);
            Path dir = Paths.get(dirPath);
            if (Files.isDirectory(dir)) {
                try (var stream = Files.list(dir)) {
                    return stream.sorted().toList();
                }
            }
            return List.of();
        });
        mockedStorageProvider = mockStatic(StorageProvider.class);
        mockedStorageProvider.when(StorageProvider::getInstance).thenReturn(spMock);

        // Replaces PowerMock's suppress of MetadatenImagesHelper.createPagination: mocked instances do nothing.
        mockedImagesHelper = mockConstruction(MetadatenImagesHelper.class);

        process = getProcess();

        File masterDir = new File(processDirectory, "images" + File.separator + "00469418X_master");
        for (int i = 1; i <= 10; i++) {
            new File(masterDir, String.format("image_%04d.tif", i)).createNewFile();
        }

        Ruleset ruleset = mock(Ruleset.class);
        when(ruleset.getDatei()).thenReturn("ruleset.xml");
        when(ruleset.getPreferences()).thenReturn(prefs);
        process.setRegelsatz(ruleset);
    }

    @AfterEach
    public void tearDown() {
        mockedConfigHelper.close();
        mockedVariableReplacer.close();
        mockedMetadatenHelper.close();
        mockedMetadataManager.close();
        mockedProcessManager.close();
        mockedHelper.close();
        mockedStorageProvider.close();
        mockedImagesHelper.close();
    }

    public Process getProcess() {
        Project project = new Project();
        project.setTitel("MetadataImportPerImageProject");

        Process process = new Process();
        process.setTitel("00469418X");
        process.setProjekt(project);
        process.setId(1);
        List<Step> steps = new ArrayList<>();
        step = new Step();
        step.setReihenfolge(1);
        step.setProzess(process);
        step.setTitel("test step");
        step.setBearbeitungsstatusEnum(StepStatus.OPEN);
        User user = new User();
        user.setVorname("Firstname");
        user.setNachname("Lastname");
        step.setBearbeitungsbenutzer(user);
        steps.add(step);

        process.setSchritte(steps);

        try {
            createProcessDirectory(processDirectory);
        } catch (IOException e) {
            throw new java.io.UncheckedIOException(e);
        }

        return process;
    }

    private void createProcessDirectory(File processDirectory) throws IOException {
        File imageDirectory = new File(processDirectory.getAbsolutePath(), "images");
        imageDirectory.mkdir();
        File masterDirectory = new File(imageDirectory.getAbsolutePath(), "00469418X_master");
        masterDirectory.mkdir();
        File mediaDirectory = new File(imageDirectory.getAbsolutePath(), "00469418X_media");
        mediaDirectory.mkdir();
    }

    /**
     * Creates test row data matching the 10 pages in meta.xml. Structure: 2 URIs, varying folders, varying captions; one empty structure value.
     */
    private List<Map<String, String>> createTestRows() {
        List<Map<String, String>> rows = new ArrayList<>();
        // uri1, folder1: pages 1-3, two captions
        rows.add(row("uri1", "folder1", "p1", "caption1"));
        rows.add(row("uri1", "folder1", "p2", "caption1"));
        rows.add(row("uri1", "folder1", "p3", "caption2"));
        // uri1, folder2: pages 4-5
        rows.add(row("uri1", "folder2", "p4", "caption3"));
        rows.add(row("uri1", "folder2", "p5", "caption3"));
        // uri2, folder3: pages 6-8, two captions
        rows.add(row("uri2", "folder3", "p6", "caption4"));
        rows.add(row("uri2", "folder3", "p7", "caption4"));
        rows.add(row("uri2", "folder3", "p8", "caption5"));
        // uri2, empty structure (→ "Unnamed folder"): pages 9-10
        rows.add(row("uri2", "", "p9", "caption6"));
        rows.add(row("uri2", "", "p10", "caption6"));
        return rows;
    }

    private Map<String, String> row(String uri, String structure, String label, String caption) {
        Map<String, String> r = new HashMap<>();
        r.put("URI", uri);
        r.put("Structure", structure);
        r.put("Label", label);
        r.put("Caption", caption);
        return r;
    }

    private void createTestExcel(File file, List<Map<String, String>> dataRows) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");
            // Header row
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("URI");
            header.createCell(1).setCellValue("Structure");
            header.createCell(2).setCellValue("Label");
            header.createCell(3).setCellValue("Caption");
            // Data rows
            for (int i = 0; i < dataRows.size(); i++) {
                Map<String, String> data = dataRows.get(i);
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(data.getOrDefault("URI", ""));
                row.createCell(1).setCellValue(data.getOrDefault("Structure", ""));
                row.createCell(2).setCellValue(data.getOrDefault("Label", ""));
                row.createCell(3).setCellValue(data.getOrDefault("Caption", ""));
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }
}
