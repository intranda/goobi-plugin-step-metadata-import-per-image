package de.intranda.goobi.plugins;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

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

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.easymock.EasyMock;
import org.goobi.beans.Process;
import org.goobi.beans.Project;
import org.goobi.beans.Ruleset;
import org.goobi.beans.Step;
import org.goobi.beans.User;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.junit.runner.RunWith;
import org.powermock.api.easymock.PowerMock;
import org.powermock.core.classloader.annotations.PowerMockIgnore;
import org.powermock.core.classloader.annotations.PrepareForTest;
import org.powermock.modules.junit4.PowerMockRunner;

import de.sub.goobi.config.ConfigurationHelper;
import de.sub.goobi.helper.VariableReplacer;
import de.sub.goobi.helper.enums.StepStatus;
import de.sub.goobi.metadaten.MetadatenHelper;
import de.sub.goobi.persistence.managers.MetadataManager;
import de.sub.goobi.persistence.managers.ProcessManager;
import ugh.dl.DocStruct;
import ugh.dl.Fileformat;
import ugh.dl.Prefs;
import ugh.fileformats.mets.MetsMods;

@RunWith(PowerMockRunner.class)
@PrepareForTest({ MetadatenHelper.class, VariableReplacer.class, ConfigurationHelper.class, ProcessManager.class,
        MetadataManager.class })
@PowerMockIgnore({ "javax.management.*", "javax.xml.*", "org.xml.*", "org.w3c.*", "javax.net.ssl.*", "jdk.internal.reflect.*" })
public class MetadataImportPerImagePluginTest {

    private static String resourcesFolder;

    @Rule
    public TemporaryFolder folder = new TemporaryFolder();

    private File processDirectory;
    private File metadataDirectory;
    private Process process;
    private Step step;
    private Prefs prefs;
    private Path metaTarget;

    @BeforeClass
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
    public void testParseExcel() throws Exception {
        // Create a test Excel file with known content
        File excelFile = folder.newFile("test.xlsx");
        createTestExcel(excelFile, createTestRows());

        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.columnUri = "URI";
        plugin.columnStructure = "Structure";
        plugin.columnLabel = "Label";
        plugin.columnCaption = "Caption";

        List<Map<String, String>> rows = plugin.parseExcel(excelFile.getAbsolutePath());

        assertEquals(10, rows.size());
        assertEquals("uri1", rows.get(0).get("URI"));
        assertEquals("folder1", rows.get(0).get("Structure"));
        assertEquals("p1", rows.get(0).get("Label"));
        assertEquals("caption1", rows.get(0).get("Caption"));
        // Row with empty structure
        assertEquals("", rows.get(8).get("Structure"));
    }

    @Test
    public void testBuildStructureCreatesHierarchy() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
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
        plugin.columnLabel = "Label";

        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        List<Map<String, String>> rows = createTestRows();
        plugin.buildStructure(ff, prefs, rows);

        List<DocStruct> pages = ff.getDigitalDocument().getPhysicalDocStruct().getAllChildren();
        assertEquals("p1", pages.get(0).getOrderLabel());
        assertEquals("p5", pages.get(4).getOrderLabel());
        assertEquals("p10", pages.get(9).getOrderLabel());
    }

    @Test
    public void testBuildStructureLinksPages() throws Exception {
        MetadataImportPerImageStepPlugin plugin = new MetadataImportPerImageStepPlugin();
        plugin.initialize(step, "something");
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

    @Before
    public void setUp() throws Exception {
        metadataDirectory = folder.newFolder("metadata");
        processDirectory = new File(metadataDirectory + File.separator + "1");
        processDirectory.mkdirs();
        String metadataDirectoryName = metadataDirectory.getAbsolutePath() + File.separator;

        Path metaSource = Paths.get(resourcesFolder, "meta.xml");
        metaTarget = Paths.get(processDirectory.getAbsolutePath(), "meta.xml");
        Files.copy(metaSource, metaTarget);

        Path anchorSource = Paths.get(resourcesFolder, "meta_anchor.xml");
        Path anchorTarget = Paths.get(processDirectory.getAbsolutePath(), "meta_anchor.xml");
        Files.copy(anchorSource, anchorTarget);

        PowerMock.mockStatic(ConfigurationHelper.class);
        ConfigurationHelper configurationHelper = EasyMock.createMock(ConfigurationHelper.class);
        EasyMock.expect(ConfigurationHelper.getInstance()).andReturn(configurationHelper).anyTimes();
        EasyMock.expect(configurationHelper.getMetsEditorLockingTime()).andReturn(1800000l).anyTimes();
        EasyMock.expect(configurationHelper.isAllowWhitespacesInFolder()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.useS3()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.isUseProxy()).andReturn(false).anyTimes();
        EasyMock.expect(configurationHelper.getGoobiContentServerTimeOut()).andReturn(60000).anyTimes();
        EasyMock.expect(configurationHelper.getMetadataFolder()).andReturn(metadataDirectoryName).anyTimes();
        EasyMock.expect(configurationHelper.getRulesetFolder()).andReturn(resourcesFolder).anyTimes();
        EasyMock.expect(configurationHelper.getProcessImagesMainDirectoryName()).andReturn("00469418X_media").anyTimes();
        EasyMock.expect(configurationHelper.isUseMasterDirectory()).andReturn(true).anyTimes();
        EasyMock.expect(configurationHelper.getConfigurationFolder()).andReturn(resourcesFolder).anyTimes();
        EasyMock.expect(configurationHelper.getNumberOfMetaBackups()).andReturn(0).anyTimes();
        EasyMock.replay(configurationHelper);

        PowerMock.mockStatic(VariableReplacer.class);
        EasyMock.expect(VariableReplacer.simpleReplace(EasyMock.anyString(), EasyMock.anyObject()))
                .andReturn("00469418X_media").anyTimes();
        PowerMock.replay(VariableReplacer.class);

        prefs = new Prefs();
        prefs.loadPrefs(resourcesFolder + "ruleset.xml");
        Fileformat ff = new MetsMods(prefs);
        ff.read(metaTarget.toString());

        PowerMock.mockStatic(MetadatenHelper.class);
        EasyMock.expect(MetadatenHelper.getMetaFileType(EasyMock.anyString())).andReturn("mets").anyTimes();
        EasyMock.expect(MetadatenHelper.getFileformatByName(EasyMock.anyString(), EasyMock.anyObject())).andReturn(ff)
                .anyTimes();
        EasyMock.expect(MetadatenHelper.getMetadataOfFileformat(EasyMock.anyObject(), EasyMock.anyBoolean()))
                .andReturn(Collections.emptyMap()).anyTimes();
        PowerMock.replay(MetadatenHelper.class);

        PowerMock.mockStatic(MetadataManager.class);
        MetadataManager.updateMetadata(1, Collections.emptyMap());
        MetadataManager.updateJSONMetadata(1, Collections.emptyMap());
        PowerMock.replay(MetadataManager.class);
        PowerMock.replay(ConfigurationHelper.class);

        process = getProcess();

        Ruleset ruleset = PowerMock.createMock(Ruleset.class);
        ruleset.setTitel("ruleset");
        ruleset.setDatei("ruleset.xml");
        EasyMock.expect(ruleset.getDatei()).andReturn("ruleset.xml").anyTimes();
        process.setRegelsatz(ruleset);
        EasyMock.expect(ruleset.getPreferences()).andReturn(prefs).anyTimes();
        PowerMock.replay(ruleset);
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
     * Creates test row data matching the 10 pages in meta.xml.
     * Structure: 2 URIs, varying folders, varying captions; one empty structure value.
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
