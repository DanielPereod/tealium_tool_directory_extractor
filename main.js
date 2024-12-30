window.audiencesCollection = window.gApp.inMemoryModels.audienceCollection.models;
  window.quantifierCollection = window.gApp.inMemoryModels.quantifierCollection.models;
  window.transformationCollection = window.gApp.inMemoryModels.transformationCollection.models;
  window.ruleCollection = window.gApp.inMemoryModels.ruleCollection.models;
  window.uniquePaths = new Set();

  function mergeCellsDown(worksheet, column, isFormula = false) {
    let previousValue = null;
    let mergeStartRow = null;

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      let cellValue;
      if (isFormula) {
        cellValue = row.getCell(column).formula;
        const match = cellValue.match(/HYPERLINK\("#Attributes!B"&MATCH\(".*",Attributes!B:B,0\),"(.*)"\)/);
        cellValue = match[1];
      } else {
        cellValue = row.getCell(column).value;
      }

      if (cellValue !== previousValue) {
        if (mergeStartRow && rowNumber - mergeStartRow > 1) {
          worksheet.mergeCells(column + mergeStartRow + ":" + column + (rowNumber - 1));
        }
        mergeStartRow = rowNumber;
        previousValue = cellValue;
      }
    });

    if (mergeStartRow && worksheet.rowCount - mergeStartRow > 0) {
      worksheet.mergeCells(column + mergeStartRow + ":" + column + worksheet.rowCount);
    }
  }

  function extractOperands(obj, operands = []) {
    for (let key in obj) {
      if (key === "operand1" || key === "operand2") {
        operands.push(obj[key]);
      } else if (typeof obj[key] === "object" && obj[key] !== null) {
        extractOperands(obj[key], operands);
      }
    }
    return operands;
  }

  function processRules(rules, path, uniquePaths) {
    for (let i = 0; i < rules.length; i++) {
      const ruleId = rules[i];
      const rule = ruleCollection.find((r) => r.attributes.id === ruleId);
      if (rule) {
        const ruleQuantifiers = extractOperands(JSON.parse(rule.attributes.logic || "{}"));
        if (ruleQuantifiers.length > 0) {
          processQuantifiers(ruleQuantifiers, [...path, rule.attributes.name], uniquePaths);
        }
      }
    }
  }

  function processQuantifiers(quantifiers, path, uniquePaths) {
    for (let j = 0; j < quantifiers.length; j++) {
      const quantifierId = quantifiers[j];

      const quantifier = quantifierCollection.find((q) => q.attributes.fullyQualifiedId === quantifierId);
      if (quantifier && !path.includes(quantifierId)) {
        const quantifierName = quantifier.attributes.name;
        const newPath = [...path, quantifierId, quantifierName];
        uniquePaths.add(JSON.stringify(newPath));

        const quantifierTransformations = quantifier.attributes.transformationIds;
        for (let k = 0; k < quantifierTransformations.length; k++) {
          const transformationId = quantifierTransformations[k];
          const transformation = transformationCollection.find((t) => t.attributes.id === transformationId);
          if (transformation) {
            const transformationName = transformation.attributes.name;
            const transformationRules = transformation.attributes.rules;

            if (transformationRules) {
              processRules(transformationRules, [...newPath, transformationName], uniquePaths);
            }

            const quantifiersQuantifiers = Object.values(transformation.attributes.action.attributes).map(el => {if (el && el.id) return el.id})
            if (quantifiersQuantifiers) {
              processQuantifiers(quantifiersQuantifiers, [...newPath, transformationName], uniquePaths);

            }
            
          }
        }
      }
    }
  }

  async function generateExcel(audiencesArray, attributeArray) {
    const workbook = new ExcelJS.Workbook();

    const audiencesWorksheet = workbook.addWorksheet("Audiences");
    const attributesWorksheet = workbook.addWorksheet("Attributes");

    // Add headers
    audiencesWorksheet.addRow(['ID', 'Name', 'Attribute', 'Description']);
    audiencesWorksheet.addRows(audiencesArray);
    attributesWorksheet.addRows(attributeArray);

    mergeCellsDown(audiencesWorksheet, "A");
    mergeCellsDown(audiencesWorksheet, "B");
    mergeCellsDown(audiencesWorksheet, "C", true);
    mergeCellsDown(audiencesWorksheet, "D"); // Add merging for description column

    mergeCellsDown(attributesWorksheet, "A");
    mergeCellsDown(attributesWorksheet, "B");
    mergeCellsDown(attributesWorksheet, "C");
    mergeCellsDown(attributesWorksheet, "D");
    mergeCellsDown(attributesWorksheet, "E");
    mergeCellsDown(attributesWorksheet, "F");
    mergeCellsDown(attributesWorksheet, "G");

    let buffer = await workbook.xlsx.writeBuffer();

    let blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    let link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "export.xlsx";
    link.click();
    URL.revokeObjectURL(link.href);
  }

  async function getTealiumTree() {
    for (let i = 0; i < audiencesCollection.length; i++) {
      const audience = audiencesCollection[i];
      const audienceId = audience.id;
      const audienceName = audience.attributes.name;
      const audienceQuantifiers = extractOperands(JSON.parse(audience.attributes.logic));

      processQuantifiers(audienceQuantifiers, [audienceId, audienceName], uniquePaths);
    }

    const audiencesArray = [];
    const attributeSet = new Set();
    const attributeArray = [];

    const uniquePathsArr = Array.from(uniquePaths);

    for (let i = 0; i < uniquePathsArr.length; i++) {
      const pathArray = JSON.parse(uniquePathsArr[i]);
      let prevPathArray;

      if (uniquePathsArr[i - 1]) {
        prevPathArray = JSON.parse(uniquePathsArr[i - 1]);
      }

      const linkFormula =
        'HYPERLINK("#Attributes!B"&MATCH("' + pathArray[3] + '",Attributes!B:B,0),"' + pathArray[3] + '")';

      const currentString = JSON.stringify([pathArray[0], pathArray[1], pathArray[2]]);
      const prevString = prevPathArray ? JSON.stringify([prevPathArray[0], prevPathArray[1], prevPathArray[2]]) : 0;

      if (currentString != prevString) {
        audiencesArray.push([
          pathArray[0], 
          pathArray[1], 
          { formula: linkFormula },
          audience.attributes.description || '' // Add description column
        ]);
      }
      attributeSet.add(JSON.stringify(pathArray.slice(2)));
    }

    attributeSet.forEach((el) => {
      attributeArray.push(JSON.parse(el));
    });

    await generateExcel(audiencesArray, attributeArray);
  }
