// ============================================================================
// FULLY FUNCTIONAL CONSOLIDATED POWER QUERY
// Output: Viz_GWP_Mat_UI_Decarb_Design_Summary
// ============================================================================

let
    // ============================================================================
    // STEP 1: LOAD ALL SOURCE TABLES ONCE
    // ============================================================================
    
    // Load all source tables at the beginning to avoid multiple loads
    Normalizer = Excel.CurrentWorkbook(){[Name="Table6"]}[Content],
    S1_History = Excel.CurrentWorkbook(){[Name="S1.History"]}[Content],
    S2_History = Excel.CurrentWorkbook(){[Name="S2.History"]}[Content],
    S3_Allocation_W = Excel.CurrentWorkbook(){[Name="S3.Allocation.W"]}[Content],
    S3_Normaliser = Excel.CurrentWorkbook(){[Name="S.3.Normaliser"]}[Content],
    Hist_S3 = Excel.CurrentWorkbook(){[Name="Hist.S3"]}[Content],
    S3_Intensity_Factor = Excel.CurrentWorkbook(){[Name="S3.Intensity.Factor"]}[Content],
    Climat_Decarb = Excel.CurrentWorkbook(){[Name="Climat_Decarb"]}[Content],
    Business_Plan = Excel.CurrentWorkbook(){[Name="Business.Plan"]}[Content],
    Mat_Library = Excel.CurrentWorkbook(){[Name="Mat.Library"]}[Content],
    Prct_Share_Materials = Excel.CurrentWorkbook(){[Name="Prct_Share_Materials"]}[Content],
    Input_Range = Excel.CurrentWorkbook(){[Name="Input_Range"]}[Content],
    Lean_Const_Library = Excel.CurrentWorkbook(){[Name="Lean.Const.Library"]}[Content],
    UI_Inputs_WF = Excel.CurrentWorkbook(){[Name="Table27"]}[Content],
    
    // ============================================================================
    // STEP 2: HELPER FUNCTIONS
    // ============================================================================
    
    // Function to unpivot year columns efficiently
    UnpivotYears = (table as table, keepColumns as list) =>
        let
            unpivoted = Table.UnpivotOtherColumns(table, keepColumns, "Year", "Value")
        in
            unpivoted,
    
    // Function to calculate percentage change from baseline
    CalculateBaselineChange = (data as table) =>
        let
            baselineValues = Table.SelectRows(data, each Text.Contains([Level], "LVL1")),
            baselineGrouped = Table.Group(baselineValues, {"ImplementingYear", "Material Category"}, {{"BaselineGWP", each List.Average([GWP.EC.FU]), type number}}),
            merged = Table.NestedJoin(data, {"ImplementingYear", "Material Category"}, baselineGrouped, {"ImplementingYear", "Material Category"}, "BaselineData", JoinKind.LeftOuter),
            expanded = Table.ExpandTableColumn(merged, "BaselineData", {"BaselineGWP"}, {"BaselineGWP"}),
            withPctChange = Table.AddColumn(expanded, "PctChangeFromBaseline", 
                each if [BaselineGWP] = null or [BaselineGWP] = 0 then null else [AvgGWP]/[BaselineGWP], type number)
        in
            withPctChange,
    
    // ============================================================================
    // STEP 3: PROCESS NORMALIZER DATA
    // ============================================================================
    
    ProcessedNormalizer = 
        let
            changedType = Table.TransformColumnTypes(Normalizer, {
                {"Business_Units", type text}, {"Unit", type text},
                {"FY14", type number}, {"FY15", type number}, {"FY16", type number}, {"FY17", type number},
                {"FY18", type number}, {"FY19", type number}, {"FY20", type number}, {"FY21", type number},
                {"FY22", type number}, {"FY23", type number}, {"FY24", type number}, {"FY25", type number},
                {"FY26", type number}, {"FY27", type number}, {"FY28", type number}, {"FY29", type number},
                {"FY30", type number}, {"FY31", type number}, {"FY32", type number}, {"FY33", type number},
                {"FY34", type number}, {"FY35", type number}, {"FY36", type number}, {"FY37", type number},
                {"FY38", type number}, {"FY39", type number}, {"FY40", type number}
            }),
            unpivoted = UnpivotYears(changedType, {"Business_Units", "Unit"}),
            renamed = Table.RenameColumns(unpivoted, {{"Value", "Normaliser_factor"}}),
            finalType = Table.TransformColumnTypes(renamed, {{"Normaliser_factor", type number}})
        in
            finalType,
    
    // ============================================================================
    // STEP 4: PROCESS SCOPE 1 & 2 DATA
    // ============================================================================
    
    ProcessedS1 = 
        let
            changedType = Table.TransformColumnTypes(S1_History, {
                {"Scope 1 (tCO2-eq)", type text}, {"FY14", Int64.Type}, {"FY15", Int64.Type},
                {"FY16", type number}, {"FY17", Int64.Type}, {"FY18", type number}, {"FY19", Int64.Type},
                {"FY20", Int64.Type}, {"FY21", Int64.Type}, {"FY22", type number}, {"FY23", type any},
                {"FY24", type any}, {"FY25", type any}, {"FY26", type any}, {"FY27", type any},
                {"FY28", type any}, {"FY29", type any}, {"FY30", type any}, {"FY31", type any},
                {"FY32", type any}, {"FY33", type any}, {"FY34", type any}, {"FY35", type any},
                {"FY36", type any}, {"FY37", type any}, {"FY38", type any}, {"FY39", type any}, {"FY40", type any}
            }),
            unpivoted = UnpivotYears(changedType, {"Scope 1 (tCO2-eq)"}),
            finalType = Table.TransformColumnTypes(unpivoted, {{"Value", type number}}),
            renamed = Table.RenameColumns(finalType, {{"Value", "Scope1"}})
        in
            renamed,
    
    ProcessedS2 = 
        let
            changedType = Table.TransformColumnTypes(S2_History, {
                {"Scope 2 (tCO2-eq)", type text}, {"FY14", Int64.Type}, {"FY15", Int64.Type},
                {"FY16", type number}, {"FY17", Int64.Type}, {"FY18", Int64.Type}, {"FY19", Int64.Type},
                {"FY20", Int64.Type}, {"FY21", Int64.Type}, {"FY22", type number}, {"FY23", type any},
                {"FY24", type any}, {"FY25", type any}, {"FY26", type any}, {"FY27", type any},
                {"FY28", type any}, {"FY29", type any}, {"FY30", type any}, {"FY31", type any},
                {"FY32", type any}, {"FY33", type any}, {"FY34", type any}, {"FY35", type any},
                {"FY36", type any}, {"FY37", type any}, {"FY38", type any}, {"FY39", type any}, {"FY40", type any}
            }),
            unpivoted = UnpivotYears(changedType, {"Scope 2 (tCO2-eq)"}),
            finalType = Table.TransformColumnTypes(unpivoted, {{"Value", type number}}),
            renamed = Table.RenameColumns(finalType, {{"Value", "Scope2"}})
        in
            renamed,
    
    // ============================================================================
    // STEP 5: PROCESS SCOPE 3 DATA
    // ============================================================================
    
    ProcessedS3Allocation = 
        let
            changedType = Table.TransformColumnTypes(S3_Allocation_W, {
                {"Region", type text}, {"Business_Units", type text}, {"GHG_Category", type text},
                {"GHG_Category_Code", type text}, {"FY20", type number}, {"FY21", type number},
                {"FY22", type number}, {"FY23", type number}, {"FY24", type number}, {"FY25", type number},
                {"FY26", type number}, {"FY27", type number}, {"FY28", type number}, {"FY29", type number},
                {"FY30", type number}, {"FY31", type number}, {"FY32", type number}, {"FY33", type number},
                {"FY34", type number}, {"FY35", type number}, {"FY36", type number}, {"FY37", type number},
                {"FY38", type number}, {"FY39", type number}, {"FY40", type number}
            }),
            unpivoted = UnpivotYears(changedType, {"Region", "Business_Units", "GHG_Category", "GHG_Category_Code"}),
            renamed = Table.RenameColumns(unpivoted, {{"Value", "Weighting_Factor"}}),
            filtered = Table.SelectRows(renamed, each [Year] <> "FY20"),
            finalType = Table.TransformColumnTypes(filtered, {{"Weighting_Factor", Percentage.Type}})
        in
            finalType,
    
    ProcessedS3Normaliser = 
        let
            changedType = Table.TransformColumnTypes(S3_Normaliser, {
                {"Region", type text}, {"Business_Units", type text}, {"Unit", type text},
                {"FY14", type number}, {"FY15", type number}, {"FY16", type number}, {"FY17", type number},
                {"FY18", type number}, {"FY19", type number}, {"FY20", type number}, {"FY21", type number},
                {"FY22", type number}, {"FY23", type number}, {"FY24", type number}, {"FY25", type number},
                {"FY26", type number}, {"FY27", type number}, {"FY28", type number}, {"FY29", type number},
                {"FY30", type number}, {"FY31", type number}, {"FY32", type number}, {"FY33", type number},
                {"FY34", type number}, {"FY35", type number}, {"FY36", type number}, {"FY37", type number},
                {"FY38", type number}, {"FY39", type number}, {"FY40", type number}
            }),
            unpivoted = UnpivotYears(changedType, {"Region", "Business_Units", "Unit"}),
            renamed = Table.RenameColumns(unpivoted, {{"Value", "Revenue"}}),
            filtered = Table.SelectRows(renamed, each not List.Contains({"FY14", "FY15", "FY16", "FY17", "FY18", "FY19", "FY20"}, [Year]))
        in
            filtered,
    
    // ============================================================================
    // STEP 6: PROCESS MATERIAL LIBRARY
    // ============================================================================
    
    ProcessedMatLibrary = 
        let
            renamed = Table.RenameColumns(Mat_Library, {{"GWP (kgCO2-e/FU)", "GWP.EC.FU"}}),
            filtered = Table.SelectRows(renamed, each Text.Lower(Text.Trim([#"Include/Exclude"])) = "in."),
            filteredLevel = Table.SelectRows(filtered, each [Level] <> null and Text.Trim([Level]) <> "" and Text.Lower(Text.Trim([Level])) <> "plz select"),
            
            // Create years list (2020-2040 for efficiency)
            yearsList = List.Numbers(2020, 21),
            
            // Get unique combinations
            uniqueCombos = Table.Distinct(filteredLevel[[Material Category], [Level]]),
            addedYears = Table.AddColumn(uniqueCombos, "ImplementingYear", each yearsList),
            expandedYears = Table.ExpandListColumn(addedYears, "ImplementingYear"),
            changedType = Table.TransformColumnTypes(expandedYears, {{"ImplementingYear", Int64.Type}}),
            
            // Merge with data
            mergedData = Table.NestedJoin(changedType, {"Material Category", "Level"}, filteredLevel, {"Material Category", "Level"}, "DataRows", JoinKind.LeftOuter),
            expandedData = Table.ExpandTableColumn(mergedData, "DataRows", {"GWP.EC.FU", "Year"}, {"GWP.EC.FU", "DataYear"}),
            filteredByYear = Table.SelectRows(expandedData, each [DataYear] <= [ImplementingYear]),
            
            // Calculate averages
            groupedAverages = Table.Group(filteredByYear, {"ImplementingYear", "Material Category", "Level"}, {{"AvgGWP", each List.Average([GWP.EC.FU]), type nullable number}}),
            
            // Calculate baseline changes
            withBaseline = CalculateBaselineChange(groupedAverages),
            
            // Add metadata
            addedCustom = Table.AddColumn(withBaseline, "S3_total-tCO2-e.GHG_Category", each "Purchased Goods and Services"),
            
            // Create decarb tag mapping
            decarbTagMapping = Table.FromRecords({
                [Material="Concrete - 40Mpa", Tag="Cement"],
                [Material="Reinforcement", Tag="Steel"],
                [Material="Structural Steel", Tag="Steel"],
                [Material="Glazing & Facade", Tag="Non-ferrous metals"],
                [Material="Other", Tag="Other"],
                [Material="Others", Tag="Other"],
                [Material="Engineered Timber", Tag="Other"]
            }),
            
            // Merge with percentage shares
            mergedWithPercentages = Table.Join(addedCustom, {"Material Category"}, Prct_Share_Materials, {"Materials"}, JoinKind.LeftOuter),
            
            // Calculate impacts
            calculatedImpacts = Table.AddColumn(mergedWithPercentages, "Mat_Initives_impacts", 
                each if [Ratio] = null then null else [Ratio] * [PctChangeFromBaseline], type number),
            
            // Merge with tags
            mergedWithTags = Table.Join(calculatedImpacts, {"Material Category"}, decarbTagMapping, {"Material"}, JoinKind.LeftOuter),
            
            // Final cleanup
            renamedDefault = Table.RenameColumns(mergedWithTags, {{"Ratio", "Prct_Share_Materials.Default"}}),
            renamedTag = Table.RenameColumns(renamedDefault, {{"Tag", "Decarb.Tag"}}),
            finalCleanup = Table.ReplaceValue(renamedTag, null, "Other", Replacer.ReplaceValue, {"Decarb.Tag"}),
            filteredYears = Table.SelectRows(finalCleanup, each [ImplementingYear] >= 2021 and [ImplementingYear] <= 2040),
            finalType = Table.TransformColumnTypes(filteredYears, {
                {"Prct_Share_Materials.Default", Percentage.Type}, 
                {"Mat_Initives_impacts", Percentage.Type}
            }),
            addedCustom1 = Table.AddColumn(finalType, "Prct_reduction_Levels", each [Prct_Share_Materials.Default] - [Mat_Initives_impacts]),
            finalType1 = Table.TransformColumnTypes(addedCustom1, {{"Prct_reduction_Levels", Percentage.Type}}),
            addedCustom2 = Table.AddColumn(finalType1, "Levels_Tag", each Text.BeforeDelimiter([Level], ":")),
            reordered = Table.ReorderColumns(addedCustom2, {
                "Material Category", "Level", "Levels_Tag", "S3_total-tCO2-e.GHG_Category", 
                "Materials", "Material", "Decarb.Tag", "ImplementingYear", "AvgGWP", "BaselineGWP", 
                "PctChangeFromBaseline", "Prct_Share_Materials.Default", "EC", "Quantity", 
                "Mat_Initives_impacts", "Prct_reduction_Levels"
            })
        in
            reordered,
    
    // ============================================================================
    // STEP 7: PROCESS CLIMATE DECARBONIZATION DATA
    // ============================================================================
    
    ProcessedClimatDecarb = 
        let
            changedType = Table.TransformColumnTypes(Climat_Decarb, {{"Year", Int64.Type}, {"Value", type number}}),
            
            // Group by all columns except Year and Value
            grouped = Table.Group(changedType, {"Model", "Scenario", "Region", "Indusry edited", "Unit"}, {
                {"AllData", each _, type table [Model=text, Scenario=text, Region=text, Industry=text, Indusry edited=text, Unit=text, Year=nullable number, Value=nullable number]}
            }),
            
            // Calculate % change from 2020 for each group
            addedCustom = Table.AddColumn(grouped, "WithPctChange", each 
                let
                    currentTable = [AllData],
                    base2020 = Table.SelectRows(currentTable, each [Year] = 2020),
                    baseValue = if Table.RowCount(base2020) > 0 then base2020[Value]{0} else null,
                    addedPct = Table.AddColumn(currentTable, "PctChange", each if baseValue = null then null else ([Value])/baseValue)
                in
                    addedPct),
            
            // Expand the tables
            expanded = Table.ExpandTableColumn(addedCustom, "WithPctChange", {"Year", "Value", "PctChange"}, {"Year", "Value", "PctChange"}),
            
            // Remove the original grouped data column
            removedColumns = Table.RemoveColumns(expanded, {"AllData"}),
            
            // Final processing
            finalType = Table.TransformColumnTypes(removedColumns, {{"PctChange", Percentage.Type}}),
            removedCols = Table.RemoveColumns(finalType, {"Unit", "Value"}),
            filtered = Table.SelectRows(removedCols, each ([Indusry edited] <> "Ammonia" and [Indusry edited] <> "Chemicals")),
            replaced = Table.ReplaceValue(filtered, "2?C", "2C", Replacer.ReplaceText, {"Scenario"}),
            finalFiltered = Table.SelectRows(replaced, each 
                ([Scenario] = "Current Policies" or [Scenario] = "Nationally Determined Contributions (NDCs)" or [Scenario] = "Net Zero 2050") and 
                ([Model] = "MESSAGEix-GLOBIOM 2.0-M-R12-NGFS"))
        in
            finalFiltered,
    
    // ============================================================================
    // STEP 8: PROCESS UI INPUTS
    // ============================================================================
    
    ProcessedUIInputs = 
        let
            filteredRows = Table.SelectRows(Input_Range, each [Column1] <> null),
            promotedHeaders = Table.PromoteHeaders(filteredRows, [PromoteAllScalars=true]),
            changedType = Table.TransformColumnTypes(promotedHeaders, {
                {"Title", type text}, {"Item", Int64.Type}, {"Current Use", type text},
                {"Comment", type text}, {"Coding.1", type text}, {"Alternative", type text},
                {"Comment_1", type text}, {"% Changes", type number}, {"Rate of changes", type text},
                {"Start Year", Int64.Type}, {"Comment_2", type text}, {"Error check", type text}
            }),
            filtered = Table.SelectRows(changedType, each ([Title] <> "")),
            removedCols = Table.RemoveColumns(filtered, {"Comment_2", "Error check"}),
            extractedText = Table.TransformColumns(removedCols, {{"Alternative", each Text.BeforeDelimiter(_, ":"), type text}}),
            removedCols1 = Table.RemoveColumns(extractedText, {"Comment_1", "Comment"}),
            addedConditional = Table.AddColumn(removedCols1, "Arrange.Alternative", each if [Alternative] = "No need" then [Coding.1] else [Alternative]),
            reordered = Table.ReorderColumns(addedConditional, {
                "Title", "Item", "Current Use", "Coding.1", "Alternative", "Arrange.Alternative", 
                "% Changes", "Rate of changes", "Start Year"
            }),
            removedCols2 = Table.RemoveColumns(reordered, {"Alternative"}),
            finalType = Table.TransformColumnTypes(removedCols2, {{"Start Year", Int64.Type}, {"% Changes", Percentage.Type}})
        in
            finalType,
    
    // ============================================================================
    // STEP 9: PROCESS LEAN CONSTRUCTION LIBRARY
    // ============================================================================
    
    ProcessedLeanConst = 
        let
            changedType = Table.TransformColumnTypes(Lean_Const_Library, {
                {"Item", Int64.Type}, {"Key", type text}, {"Strategy", type text}, {"Sub_Strategy", type text},
                {"Key2", type text}, {"Avg reduction", type number}, {"Impacts", type number}, {"Node", type any},
                {"Default", type any}, {"Comments", type any}
            }),
            filtered = Table.SelectRows(changedType, each [Item] <> null),
            removedCols = Table.RemoveColumns(filtered, {"Default", "Comments"}),
            finalType = Table.TransformColumnTypes(removedCols, {{"Avg reduction", Percentage.Type}, {"Impacts", Percentage.Type}})
        in
            finalType,
    
    // ============================================================================
    // STEP 10: PROCESS UI INPUTS WF
    // ============================================================================
    
    ProcessedUIInputsWF = 
        let
            changedType = Table.TransformColumnTypes(UI_Inputs_WF, {
                {"Region", type text}, {"BUs", type text}, {"Start Year", Int64.Type}, 
                {"End Year", Int64.Type}, {"Climate Scenario", type text}
            })
        in
            changedType,
    
    // ============================================================================
    // STEP 11: BUILD COMPLETE PIPELINE
    // ============================================================================
    
    // Start with S3 reference point
    S3RefPoint = 
        let
            // Process S3 allocation with weighting
            s3Allocation = ProcessedS3Allocation,
            
            // Merge with S3 normaliser
            mergedWithNormaliser = Table.NestedJoin(s3Allocation, {"Region", "Business_Units", "Year"}, 
                ProcessedS3Normaliser, {"Region", "Business_Units", "Year"}, "S3_Normaliser", JoinKind.LeftOuter),
            expandedNormaliser = Table.ExpandTableColumn(mergedWithNormaliser, "S3_Normaliser", 
                {"Unit", "Revenue"}, {"Unit", "Revenue"}),
            
            // Calculate normalized revenue
            addedCustom = Table.AddColumn(expandedNormaliser, "Normalised_Revenue", each [Revenue] * [Weighting_Factor]),
            removedCols = Table.RemoveColumns(addedCustom, {"Revenue", "Weighting_Factor"}),
            
            // Merge with S3 intensity factor
            mergedWithIntensity = Table.NestedJoin(removedCols, {"GHG_Category_Code", "Region", "Business_Units"}, 
                S3_Intensity_Factor, {"GHG_Category_Code", "Region", "Business_Units"}, "S3_Intensity_Factor", JoinKind.LeftOuter),
            expandedIntensity = Table.ExpandTableColumn(mergedWithIntensity, "S3_Intensity_Factor", 
                {"S3_IntensityRatio", "Unit"}, {"S3_IntensityRatio", "Unit"}),
            
            // Calculate projected S3 tCO2-e
            addedCustom1 = Table.AddColumn(expandedIntensity, "Projected_S3_tCO2e", each [Normalised_Revenue] * [S3_IntensityRatio]),
            filtered = Table.SelectRows(addedCustom1, each [Projected_S3_tCO2e] <> null)
        in
            filtered,
    
    // Connect materials to business units and GHGs
    ConnectedMaterials = 
        let
            // Merge S3 ref point with material library
            merged = Table.NestedJoin(S3RefPoint, {"GHG_Category"}, ProcessedMatLibrary, 
                {"S3_total-tCO2-e.GHG_Category"}, "Mat_Library", JoinKind.LeftOuter),
            expanded = Table.ExpandTableColumn(merged, "Mat_Library", {
                "ImplementingYear", "Material Category", "Level", "Mat_Initives_impacts", "Decarb.Tag"
            }, {"ImplementingYear", "Material Category", "Level", "Mat_Initives_impacts", "Decarb.Tag"}),
            
            // Filter for baseline level and 2025 implementation
            filtered = Table.SelectRows(expanded, each 
                ([ImplementingYear] = null or [ImplementingYear] = 2025) and 
                ([Level] = null or [Level] = "LVL1: Baseline")),
            
            // Calculate material GWP impacts
            addedCustom = Table.AddColumn(filtered, "Mat_GWP_Perct_Impacts", each [Projected_S3_tCO2e] * [Mat_Initives_impacts]),
            addedConditional = Table.AddColumn(addedCustom, "tCO2_e_fill_Gaps", each 
                if [Mat_GWP_Perct_Impacts] = null then [Projected_S3_tCO2e] else [Mat_GWP_Perct_Impacts]),
            
            // Add decarb tags
            addedConditional1 = Table.AddColumn(addedConditional, "Decarb_Tag_2", each 
                if [Decarb.Tag] = null then "O.N.M" else [Decarb.Tag]),
            addedConditional2 = Table.AddColumn(addedConditional1, "Decarb_Tag_3", each 
                if [Decarb.Tag] = null then "Other" else [Decarb.Tag]),
            
            // Add full year
            addedFullYear = Table.AddColumn(addedConditional2, "FullYear", each 
                let
                    yearText = Text.End(Text.From([Year]), 2),
                    fullYear = "20" & yearText
                in
                    fullYear, type text),
            
            finalType = Table.TransformColumnTypes(addedFullYear, {{"FullYear", Int64.Type}}),
            filteredYears = Table.SelectRows(finalType, each [Year] <> "FY20")
        in
            filteredYears,
    
    // Apply UI inputs for materials
    MaterialsWithUI = 
        let
            // Merge with UI inputs analysis
            merged = Table.NestedJoin(ConnectedMaterials, {"GHG_Category", "Decarb_Tag_3", "FullYear"}, 
                #"UI_Inputs (analysis)_Mat_Innitivies", {"GHG_Category", "Tag", "Years"}, 
                "UI_Inputs (analysis)_Mat_Innitivies", JoinKind.LeftOuter),
            
            // Expand and calculate
            expanded = Table.ExpandTableColumn(merged, "UI_Inputs (analysis)_Mat_Innitivies", 
                {"Arrange.Alternative", "Years", "Reduction impacts"}, {"Arrange.Alternative", "Years", "Reduction impacts"}),
            
            // Calculate reduction impacts
            addedConditional = Table.AddColumn(expanded, "Reduction_impacts_fix", each 
                if [Reduction impacts] = null then 1 else [Reduction impacts]),
            addedCustom1 = Table.AddColumn(addedConditional, "GWP_UI", each [tCO2_e_fill_Gaps] * [Reduction_impacts_fix]),
            addedCustom2 = Table.AddColumn(addedCustom1, "Reductions_GWP_UI", each [GWP_UI] - [tCO2_e_fill_Gaps]),
            
            // Group and calculate totals
            grouped = Table.Group(addedCustom2, {
                "Region", "Business_Units", "GHG_Category", "GHG_Category_Code", "FullYear", 
                "Projected_S3_tCO2e", "Material Category", "Decarb_Tag_3", "tCO2_e_fill_Gaps", "Year"
            }, {{"Red_GWP_UI", each List.Sum([Reductions_GWP_UI]), type number}}),
            
            addedCustom3 = Table.AddColumn(grouped, "Overall_GWP_UI", each [tCO2_e_fill_Gaps] + [Red_GWP_UI])
        in
            addedCustom3,
    
    // Apply climate decarbonization
    FinalWithClimate = 
        let
            // Merge with climate data
            merged = Table.NestedJoin(MaterialsWithUI, {"FullYear", "Decarb_Tag_3"}, 
                ProcessedClimatDecarb, {"Year", "Indusry edited"}, "Climat_Decarb", JoinKind.LeftOuter),
            expanded = Table.ExpandTableColumn(merged, "Climat_Decarb", 
                {"Model", "Scenario", "PctChange"}, {"Model", "Scenario", "PctChange"}),
            
            // Calculate final GWP
            addedCustom = Table.AddColumn(expanded, "Overall_GWP_Decarb_TCO2e", each [Overall_GWP_UI] * [PctChange]),
            
            // Group by model, scenario, and year
            grouped = Table.Group(addedCustom, {"Model", "Scenario", "FullYear"}, {
                {"Sum_decarb_GWP_T_UI", each List.Sum([Overall_GWP_Decarb_TCO2e]), type number}
            })
        in
            grouped,
    
    // ============================================================================
    // STEP 12: FINAL OUTPUT - Viz_GWP_Mat_UI_Decarb_Design_Summary
    // ============================================================================
    
    FinalOutput = 
        let
            // Merge with design initiatives
            mergedWithDesign = Table.NestedJoin(FinalWithClimate, {"FullYear"}, 
                #"Design&UI", {"Years"}, "Design_UI", JoinKind.LeftOuter),
            expandedDesign = Table.ExpandTableColumn(mergedWithDesign, "Design_UI", 
                {"Coding.4", "Reduction impacts"}, {"Coding.4", "Reduction impacts"}),
            
            // Calculate final reductions
            addedCustom = Table.AddColumn(expandedDesign, "Reduction", each [Sum_decarb_GWP_T_UI] * [Reduction impacts]),
            
            // Group and calculate final values
            grouped = Table.Group(addedCustom, {
                "Model", "Scenario", "FullYear", "Sum_decarb_GWP_T_UI"
            }, {{"Sum_Reduction", each List.Sum([Reduction]), type number}}),
            
            addedCustom1 = Table.AddColumn(grouped, "Mat_Des_Inn_GWP_Decarb", each [Sum_decarb_GWP_T_UI] - [Sum_Reduction]),
            
            // Final conditional column
            finalConditional = Table.AddColumn(addedCustom1, "Final_GWP_Decarb", each 
                if [Mat_Des_Inn_GWP_Decarb] = null then [Sum_decarb_GWP_T_UI] else [Mat_Des_Inn_GWP_Decarb]),
            
            // Clean up columns
            removedCols = Table.RemoveColumns(finalConditional, {"Sum_Reduction"}),
            renamed = Table.RenameColumns(removedCols, {
                {"Sum_decarb_GWP_T_UI", "BAU_GWP_Model"},
                {"Final_GWP_Decarb", "Mat_Des_Inn_GWP_Decarb"}
            })
        in
            renamed

in
    FinalOutput