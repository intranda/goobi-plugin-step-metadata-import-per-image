package de.intranda.goobi.plugins;

/**
 * This file is part of a plugin for Goobi - a Workflow tool for the support of mass digitization.
 *
 * Visit the websites for more information.
 *          - https://goobi.io
 *          - https://www.intranda.com
 *          - https://github.com/intranda/goobi
 *
 * This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free
 * Software Foundation; either version 2 of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
 * FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59
 * Temple Place, Suite 330, Boston, MA 02111-1307 USA
 *
 */

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.SubnodeConfiguration;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.goobi.beans.Process;
import org.goobi.beans.Step;
import org.goobi.io.BackupFileManager;
import org.goobi.production.enums.PluginGuiType;
import org.goobi.production.enums.PluginReturnValue;
import org.goobi.production.enums.PluginType;
import org.goobi.production.enums.StepReturnValue;
import org.goobi.production.plugin.interfaces.IStepPluginVersion2;

import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.config.ConfigurationHelper;
import de.sub.goobi.helper.Helper;
import de.sub.goobi.helper.StorageProvider;
import de.sub.goobi.helper.VariableReplacer;
import de.sub.goobi.helper.exceptions.DAOException;
import de.sub.goobi.helper.exceptions.SwapException;
import de.sub.goobi.metadaten.MetadatenImagesHelper;
import lombok.Getter;
import lombok.extern.log4j.Log4j2;
import org.goobi.production.enums.LogType;
import net.xeoh.plugins.base.annotations.PluginImplementation;
import ugh.dl.DigitalDocument;
import ugh.dl.DocStruct;
import ugh.dl.DocStructType;
import ugh.dl.Fileformat;
import ugh.dl.Metadata;
import ugh.dl.MetadataType;
import ugh.dl.Prefs;
import ugh.dl.Reference;
import ugh.exceptions.*;

@PluginImplementation
@Log4j2
public class MetadataImportPerImageStepPlugin implements IStepPluginVersion2 {
    @Getter
    private String title = "intranda_step_metadata_import_per_image";
    @Getter
    private Step step;
    private String returnPath;

    // Configuration (package-private for testing)
    String excelFilePath;
    String columnLabel;
    String paginationLabelMetadata = "logicalPageNumber";
    List<HierarchyLevel> hierarchyLevels;

    @Override
    public void initialize(Step step, String returnPath) {
        this.returnPath = returnPath;
        this.step = step;

        SubnodeConfiguration myconfig = ConfigPlugins.getProjectAndStepConfig(title, step);
        excelFilePath = myconfig.getString("excelFile", "");
        columnLabel = myconfig.getString("columnLabel", "Label");
        paginationLabelMetadata = myconfig.getString("paginationLabelMetadata", "logicalPageNumber");

        hierarchyLevels = new ArrayList<>();
        List<HierarchicalConfiguration> levelConfigs = myconfig.configurationAt("hierarchy").configurationsAt("level");
        for (HierarchicalConfiguration levelConfig : levelConfigs) {
            HierarchyLevel level = new HierarchyLevel();
            level.structType = levelConfig.getString("@structType", "");
            level.groupByColumn = levelConfig.getString("@groupByColumn", "");
            level.metadataField = levelConfig.getString("@metadataField", "");
            level.fallbackTitle = levelConfig.getString("@fallbackTitle", "");
            hierarchyLevels.add(level);
        }

        log.info("MetadataImportPerImage step plugin initialized");
    }

    @Override
    public PluginGuiType getPluginGuiType() {
        return PluginGuiType.NONE;
    }

    @Override
    public String getPagePath() {
        return null;
    }

    @Override
    public PluginType getType() {
        return PluginType.Step;
    }

    @Override
    public String cancel() {
        return "/uii" + returnPath;
    }

    @Override
    public String finish() {
        return "/uii" + returnPath;
    }

    @Override
    public int getInterfaceVersion() {
        return 0;
    }

    @Override
    public HashMap<String, StepReturnValue> validate() {
        return null;
    }

    @Override
    public boolean execute() {
        PluginReturnValue ret = run();
        return ret != PluginReturnValue.ERROR;
    }

    @Override
    public PluginReturnValue run() {
        Process process = step.getProzess();
        log.info("Running MetadataImportPerImage plugin for process {}", process.getTitel());

        // Read metadata file — used for variable replacement and structure building
        Fileformat fileformat;
        Prefs prefs;
        DigitalDocument dd;
        try {
            fileformat = process.readMetadataFile();
            prefs = process.getRegelsatz().getPreferences();
            dd = fileformat.getDigitalDocument();
        } catch (Exception e) {
            return reportError(process, "Failed to read metadata file: " + e.getMessage());
        }

        // Validate configuration against ruleset
        ValidationResult configValidation = validateConfig(prefs);
        if (configValidation.hasErrors()) {
            for (String error : configValidation.errors) {
                Helper.addMessageToProcessJournal(process.getId(), LogType.ERROR, error);
            }
            return reportError(process, String.join("; ", configValidation.errors));
        }

        // Ensure pagination structures exist (idempotent — safe to call even when pages already exist)
        MetadatenImagesHelper mih = new MetadatenImagesHelper(prefs, dd);
        try {
            mih.createPagination(process, null);
        } catch (TypeNotAllowedForParentException | IOException | SwapException | DAOException e) {
            return reportError(process, "Failed to create pagination: " + e.getMessage());
        }

        // Resolve Excel file path (supports Goobi folder variables like {folder.images.import})
        VariableReplacer vr = new VariableReplacer(dd, prefs, process, step);
        String resolvedExcelPath = vr.replace(excelFilePath);
        log.debug("Excel file path resolved to: {}", resolvedExcelPath);

        // Parse Excel rows
        ParseResult parseResult;
        try {
            parseResult = parseExcel(resolvedExcelPath);
        } catch (IOException e) {
            return reportError(process, "Failed to read Excel file " + resolvedExcelPath + ": " + e.getMessage());
        }

        // Get images from master folder
        List<java.nio.file.Path> images;
        try {
            String masterFolder = process.getImagesOrigDirectory(false);
            images = StorageProvider.getInstance().listFiles(masterFolder).stream()
                    .filter(p -> !Files.isDirectory(p))
                    .filter(p -> !p.getFileName().toString().startsWith("."))
                    .sorted()
                    .collect(Collectors.toList());
        } catch (Exception e) {
            return reportError(process, "Failed to list master folder images: " + e.getMessage());
        }

        // Count physical pages in metadata
        DocStruct physRoot = dd.getPhysicalDocStruct();
        int physPageCount = (physRoot != null && physRoot.getAllChildren() != null) ? physRoot.getAllChildren().size() : 0;

        // Validate Excel data
        ValidationResult dataValidation = validateExcelData(parseResult.rows, prefs, images.size(), physPageCount,
                parseResult.sheetCount);

        // Merge parse warnings into data validation warnings
        for (String warning : parseResult.parseWarnings) {
            dataValidation.addWarning(warning);
        }

        if (dataValidation.hasErrors()) {
            for (String error : dataValidation.errors) {
                Helper.addMessageToProcessJournal(process.getId(), LogType.ERROR, error);
            }
            return reportError(process, String.join("; ", dataValidation.errors));
        }

        // Log warnings to journal but continue processing
        for (String warning : dataValidation.warnings) {
            log.warn(warning);
            Helper.addMessageToProcessJournal(process.getId(), LogType.WARN, warning);
        }

        // Create backup of meta.xml before modifying
        try {
            String processFolder = Paths.get(process.getMetadataFilePath()).getParent().toString();
            int backupCount = ConfigurationHelper.getInstance().getNumberOfMetaBackups();
            if (backupCount > 0) {
                BackupFileManager.createBackup(processFolder, "meta.xml", backupCount, false);
            }
        } catch (Exception e) {
            return reportError(process, "Failed to create meta.xml backup: " + e.getMessage());
        }

        // Build logical structure from Excel data
        try {
            buildStructure(fileformat, prefs, parseResult.rows);
        } catch (Exception e) {
            return reportError(process, "Failed to build metadata structure: " + e.getMessage());
        }

        // Write updated metadata file
        try {
            process.writeMetadataFile(fileformat);
        } catch (Exception e) {
            return reportError(process, "Failed to write metadata file: " + e.getMessage());
        }

        log.info("MetadataImportPerImage plugin completed successfully for process {}", process.getTitel());
        return PluginReturnValue.FINISH;
    }

    private PluginReturnValue reportError(Process process, String message) {
        log.error(message);
        Helper.setFehlerMeldungUntranslated(message);
        Helper.addMessageToProcessJournal(process.getId(), LogType.ERROR, message);
        return PluginReturnValue.ERROR;
    }

    /**
     * Validates the plugin configuration against the given ruleset.
     * Checks that hierarchy levels are defined and that their struct types
     * and metadata fields exist in the ruleset.
     */
    ValidationResult validateConfig(Prefs prefs) {
        ValidationResult result = new ValidationResult();

        if (hierarchyLevels == null || hierarchyLevels.isEmpty()) {
            result.addError("No hierarchy levels configured");
            return result;
        }

        for (HierarchyLevel level : hierarchyLevels) {
            if (prefs.getDocStrctTypeByName(level.structType) == null) {
                result.addError("Struct type not found in ruleset: " + level.structType);
            }
            if (StringUtils.isNotBlank(level.metadataField)
                    && prefs.getMetadataTypeByName(level.metadataField) == null) {
                result.addError("Metadata field not found in ruleset: " + level.metadataField);
            }
        }

        return result;
    }

    /**
     * Validates the parsed Excel data against expected constraints.
     * Collects all errors and warnings rather than failing on the first problem.
     */
    ValidationResult validateExcelData(List<Map<String, String>> rows, Prefs prefs, int imageCount, int physPageCount,
            int sheetCount) {
        ValidationResult result = new ValidationResult();

        // Check 4: empty Excel (0 data rows)
        if (rows.isEmpty()) {
            result.addError("Excel file contains no data rows");
            return result;
        }

        if (rows.size() != imageCount) {
            result.addError("Row count mismatch: Excel has " + rows.size()
                    + " data rows, master folder has " + imageCount + " images");
        }

        if (rows.size() != physPageCount) {
            result.addError("Row count mismatch: Excel has " + rows.size()
                    + " data rows, metadata has " + physPageCount + " physical pages");
        }

        // Check 1: label column header missing
        if (!rows.get(0).containsKey(columnLabel)) {
            result.addError("Label column '" + columnLabel + "' not found in Excel headers");
        }

        // Check missing grouping columns
        for (HierarchyLevel level : hierarchyLevels) {
            if (StringUtils.isNotBlank(level.groupByColumn) && !rows.get(0).containsKey(level.groupByColumn)) {
                result.addError("Missing column in Excel: " + level.groupByColumn);
            }
        }

        // Collect hierarchy levels that have a non-blank groupByColumn
        List<HierarchyLevel> groupingLevels = new ArrayList<>();
        for (HierarchyLevel level : hierarchyLevels) {
            if (StringUtils.isNotBlank(level.groupByColumn)) {
                groupingLevels.add(level);
            }
        }

        // Per-row checks
        // Track values seen and last value per groupByColumn for non-contiguous detection (check 7)
        Map<String, Set<String>> seenValues = new HashMap<>();
        Map<String, String> lastValues = new HashMap<>();
        for (HierarchyLevel level : groupingLevels) {
            seenValues.put(level.groupByColumn, new HashSet<>());
            lastValues.put(level.groupByColumn, null);
        }

        for (int rowIdx = 0; rowIdx < rows.size(); rowIdx++) {
            Map<String, String> row = rows.get(rowIdx);
            int excelRowNumber = rowIdx + 2; // 1-based, accounting for header

            // Check 3: empty groupByColumn value with no fallback
            for (HierarchyLevel level : groupingLevels) {
                String value = row.getOrDefault(level.groupByColumn, "").trim();
                if (value.isEmpty() && StringUtils.isBlank(level.fallbackTitle)) {
                    result.addError("Row " + excelRowNumber + ": empty value in column '" + level.groupByColumn
                            + "' with no fallback configured");
                }
            }

            // Check 6: all grouping columns empty
            if (!groupingLevels.isEmpty()) {
                boolean allEmpty = true;
                for (HierarchyLevel level : groupingLevels) {
                    String value = row.getOrDefault(level.groupByColumn, "").trim();
                    if (!value.isEmpty()) {
                        allEmpty = false;
                        break;
                    }
                }
                if (allEmpty) {
                    result.addError("Row " + excelRowNumber + ": all grouping columns are empty");
                }
            }

            // Check 5: XML-invalid control characters in any cell
            for (Map.Entry<String, String> entry : row.entrySet()) {
                if (containsInvalidXmlChars(entry.getValue())) {
                    result.addError("Row " + excelRowNumber + ", column '" + entry.getKey()
                            + "': contains invalid XML control characters");
                }
            }

            // Check 7: non-contiguous duplicate keys
            for (HierarchyLevel level : groupingLevels) {
                String value = row.getOrDefault(level.groupByColumn, "").trim();
                if (!value.isEmpty()) {
                    Set<String> seen = seenValues.get(level.groupByColumn);
                    String lastValue = lastValues.get(level.groupByColumn);
                    if (seen.contains(value) && !value.equals(lastValue)) {
                        result.addWarning("Non-contiguous values in column '" + level.groupByColumn
                                + "': value '" + value + "' reappears at row " + excelRowNumber);
                    }
                    seen.add(value);
                    lastValues.put(level.groupByColumn, value);
                }
            }
        }

        // Check 8: multiple sheets
        if (sheetCount > 1) {
            result.addWarning("Workbook contains " + sheetCount + " sheets; only the first sheet is processed");
        }

        return result;
    }

    /**
     * Checks whether a string contains XML-invalid control characters
     * (characters below 0x20 except tab, newline, and carriage return).
     */
    private static boolean containsInvalidXmlChars(String value) {
        if (value == null) {
            return false;
        }
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            if (c < 0x20 && c != '\t' && c != '\n' && c != '\r') {
                return true;
            }
        }
        return false;
    }

    /**
     * Reads all data rows from the first sheet of the given Excel file.
     * The first row is treated as the header. Each subsequent non-empty row becomes a map of column name to cell value.
     *
     * @return a ParseResult containing the rows, sheet count, and any parse warnings
     */
    ParseResult parseExcel(String filePath) throws IOException {
        List<Map<String, String>> rows = new ArrayList<>();
        List<String> parseWarnings = new ArrayList<>();
        int sheetCount;
        try (Workbook workbook = WorkbookFactory.create(Paths.get(filePath).toFile())) {
            sheetCount = workbook.getNumberOfSheets();
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                throw new IOException("Excel file has no header row: " + filePath);
            }

            // Map column header to column index, detecting duplicates
            Map<String, Integer> columnIndex = new LinkedHashMap<>();
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                Cell cell = headerRow.getCell(c);
                if (cell != null) {
                    String header = getCellValue(cell).trim();
                    if (!header.isEmpty()) {
                        if (columnIndex.containsKey(header)) {
                            throw new IOException("Duplicate column header: " + header);
                        }
                        columnIndex.put(header, c);
                    }
                }
            }

            // Read data rows, skipping fully empty rows
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                boolean hasContent = false;
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    if (!getCellValue(row.getCell(c)).trim().isEmpty()) {
                        hasContent = true;
                        break;
                    }
                }
                if (!hasContent) {
                    continue;
                }

                Map<String, String> rowData = new HashMap<>();
                for (Map.Entry<String, Integer> entry : columnIndex.entrySet()) {
                    rowData.put(entry.getKey(), getCellValue(row.getCell(entry.getValue())).trim());
                }
                rows.add(rowData);
            }
        }
        return new ParseResult(rows, sheetCount, parseWarnings);
    }

    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        CellType type = cell.getCellType();
        if (type == CellType.ERROR) {
            return "";
        }
        if (type == CellType.NUMERIC) {
            double d = cell.getNumericCellValue();
            if (d == Math.floor(d) && !Double.isInfinite(d)) {
                return String.valueOf((long) d);
            }
            return String.valueOf(d);
        }
        if (type == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        }
        if (type == CellType.FORMULA) {
            CellType cachedType = cell.getCachedFormulaResultType();
            if (cachedType == CellType.ERROR) {
                return "";
            }
            try {
                return cell.getStringCellValue();
            } catch (IllegalStateException e) {
                double d = cell.getNumericCellValue();
                if (d == Math.floor(d) && !Double.isInfinite(d)) {
                    return String.valueOf((long) d);
                }
                return String.valueOf(d);
            }
        }
        String val = cell.getStringCellValue();
        return val != null ? val : "";
    }

    /**
     * Rebuilds the logical structure of the metadata document based on the Excel rows.
     * Removes all existing children of the content root, then creates new hierarchy
     * elements according to the configured levels. Sets ORDERLABEL on physical pages
     * from the label column. Links each page to its containing logical elements.
     */
    void buildStructure(Fileformat fileformat, Prefs prefs, List<Map<String, String>> rows)
            throws TypeNotAllowedAsChildException, MetadataTypeNotAllowedException,
            TypeNotAllowedForParentException, PreferencesException {
        DigitalDocument document = fileformat.getDigitalDocument();
        DocStruct contentRoot = getContentRoot(document);
        if (contentRoot == null) {
            throw new IllegalStateException("No content root found in logical structure");
        }

        final MetadataType logicalPageNumberType = prefs.getMetadataTypeByName(paginationLabelMetadata);
        if (logicalPageNumberType == null) {
            log.warn("Metadata type '{}' not found in ruleset; skipping pagination labels", paginationLabelMetadata);
        }

        // Remove existing children of content root (clean up references first)
        if (contentRoot.getAllChildren() != null) {
            for (DocStruct child : new ArrayList<>(contentRoot.getAllChildren())) {
                removeReferencesRecursively(child);
                contentRoot.removeChild(child);
            }
        }

        // Remove existing direct page references from content root
        if (contentRoot.getAllToReferences() != null) {
            for (Reference ref : new ArrayList<>(contentRoot.getAllToReferences())) {
                contentRoot.removeReferenceTo(ref.getTarget());
            }
        }

        // Collect physical pages in document order
        List<DocStruct> physicalPages = new ArrayList<>();
        DocStruct physicalRoot = document.getPhysicalDocStruct();
        if (physicalRoot != null && physicalRoot.getAllChildren() != null) {
            physicalPages.addAll(physicalRoot.getAllChildren());
        }
        if (physicalPages.size() < rows.size()) {
            throw new IllegalStateException("Not enough physical pages (" + physicalPages.size()
                    + ") for " + rows.size() + " data rows");
        }

        int numLevels = hierarchyLevels.size();
        String[] currentKeys = new String[numLevels];
        DocStruct[] currentStructs = new DocStruct[numLevels];

        for (int rowIdx = 0; rowIdx < rows.size(); rowIdx++) {
            Map<String, String> row = rows.get(rowIdx);
            DocStruct page = physicalPages.get(rowIdx);

            // Set pagination label on physical page
            String label = row.getOrDefault(columnLabel, "");
            if (logicalPageNumberType != null && !label.isEmpty()) {
                List<? extends Metadata> pageLabelMetadata = page.getAllMetadataByType(logicalPageNumberType);
                if (!pageLabelMetadata.isEmpty()) {
                    pageLabelMetadata.getFirst().setValue(label);
                } else {
                    Metadata pageLabel = new Metadata(logicalPageNumberType);
                    pageLabel.setValue(label);
                    page.addMetadata(pageLabel);
                }
            }

            // Find which level changed first (null currentKeys[lvl] triggers change)
            int changedAt = numLevels;
            for (int lvl = 0; lvl < numLevels; lvl++) {
                String newKey = getKey(row, lvl);
                if (!newKey.equals(currentKeys[lvl])) {
                    changedAt = lvl;
                    break;
                }
            }

            // Create new DocStruct elements from changedAt level downward
            for (int lvl = changedAt; lvl < numLevels; lvl++) {
                String key = getKey(row, lvl);
                HierarchyLevel levelConfig = hierarchyLevels.get(lvl);

                DocStructType structType = prefs.getDocStrctTypeByName(levelConfig.structType);
                if (structType == null) {
                    throw new IllegalStateException("DocStruct type not found in ruleset: " + levelConfig.structType);
                }
                DocStruct newStruct = document.createDocStruct(structType);

                if (lvl == 0) {
                    contentRoot.addChild(newStruct);
                } else {
                    currentStructs[lvl - 1].addChild(newStruct);
                }

                if (StringUtils.isNotBlank(levelConfig.metadataField)) {
                    MetadataType mdType = prefs.getMetadataTypeByName(levelConfig.metadataField);
                    if (mdType != null) {
                        Metadata md = new Metadata(mdType);
                        md.setValue(key);
                        try {
                            newStruct.addMetadata(md);
                        } catch (MetadataTypeNotAllowedException e) {
                            log.warn("Cannot add metadata '{}' to '{}': {}", levelConfig.metadataField,
                                    levelConfig.structType, e.getMessage());
                        }
                    }
                }

                currentKeys[lvl] = key;
                currentStructs[lvl] = newStruct;
            }

            // Link this page to the content root and all active hierarchy levels
            contentRoot.addReferenceTo(page, "logical_physical");
            for (int lvl = 0; lvl < numLevels; lvl++) {
                if (currentStructs[lvl] != null) {
                    currentStructs[lvl].addReferenceTo(page, "logical_physical");
                }
            }
        }
    }

    /**
     * Returns the logical DocStruct that represents the current process's content.
     * If the logical root is an anchor (multi-volume parent), returns its first child instead.
     */
    private DocStruct getContentRoot(DigitalDocument document) {
        DocStruct root = document.getLogicalDocStruct();
        if (root != null && root.getType().isAnchor()) {
            List<DocStruct> children = root.getAllChildren();
            if (children != null && !children.isEmpty()) {
                return children.get(0);
            }
            throw new IllegalStateException("Anchor document has no child volumes");
        }
        return root;
    }

    /**
     * Recursively removes all TO-references from the given DocStruct and its children.
     * This ensures that physical pages no longer hold back-references to structures
     * that are about to be removed.
     */
    private void removeReferencesRecursively(DocStruct node) {
        if (node.getAllToReferences() != null) {
            for (Reference ref : new ArrayList<>(node.getAllToReferences())) {
                node.removeReferenceTo(ref.getTarget());
            }
        }
        if (node.getAllChildren() != null) {
            for (DocStruct child : node.getAllChildren()) {
                removeReferencesRecursively(child);
            }
        }
    }

    /**
     * Returns the grouping key for the given hierarchy level from the current row.
     * Applies the configured fallback title when the column value is empty.
     */
    private String getKey(Map<String, String> row, int levelIndex) {
        HierarchyLevel levelConfig = hierarchyLevels.get(levelIndex);
        String value = row.getOrDefault(levelConfig.groupByColumn, "").trim();
        if (value.isEmpty() && StringUtils.isNotBlank(levelConfig.fallbackTitle)) {
            value = levelConfig.fallbackTitle;
        }
        return value;
    }

    static class HierarchyLevel {
        String structType;
        String groupByColumn;
        String metadataField;
        String fallbackTitle;
    }

    static class ValidationResult {
        final List<String> errors = new ArrayList<>();
        final List<String> warnings = new ArrayList<>();

        void addError(String msg) {
            errors.add(msg);
        }

        void addWarning(String msg) {
            warnings.add(msg);
        }

        boolean hasErrors() {
            return !errors.isEmpty();
        }
    }

    static class ParseResult {
        final List<Map<String, String>> rows;
        final int sheetCount;
        final List<String> parseWarnings;

        ParseResult(List<Map<String, String>> rows, int sheetCount, List<String> parseWarnings) {
            this.rows = rows;
            this.sheetCount = sheetCount;
            this.parseWarnings = parseWarnings;
        }
    }
}
