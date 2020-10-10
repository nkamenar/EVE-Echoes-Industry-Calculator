/**
 * Copyright 2020 Nathan Kamenar
 *
 * This file is subject to the terms and conditions defined in file 'LICENSE.md',
 * which can be found at https://github.com/nkamenar/EVE-Echoes-Industry-Calculator.
 */

if (!String.prototype.newStringFunc) {
	String.prototype.toFormattedName = function () {
		return this.toLowerCase().replace(' ', '');
	};
}

const minerals = {
	tritanium: {
		current: 0,
		needed: 0,
	},
	pyerite: {
		current: 0,
		needed: 0,
	},
	mexallon: {
		current: 0,
		needed: 0,
	},
	isogen: {
		current: 0,
		needed: 0,
	},
	nocxium: {
		current: 0,
		needed: 0,
	},
	zydrine: {
		current: 0,
		needed: 0,
	},
	megacyte: {
		current: 0,
		needed: 0,
	},
	morphite: {
		current: 0,
		needed: 0,
	},
};

const ActiveSpreadsheet = SpreadsheetApp.getActive();

// eslint-disable-next-line no-unused-vars
function onEdit(e) {
	try {
		const componentNameCell = ActiveSpreadsheet.getRangeByName('selectedComponentName').getCell(1, 1).getA1Notation();
		const reprocessingEfficiencyCells = getRangeA1Values(ActiveSpreadsheet.getRangeByName('oreReprocessingEfficiencyValues'));
		const oreToggles = getRangeA1Values(ActiveSpreadsheet.getRangeByName('oreToggles'));

		const optimizeForCell = ActiveSpreadsheet.getRangeByName('optimizeFor').getCell(1, 1).getA1Notation();
		const materialEfficiencyCell = ActiveSpreadsheet.getRangeByName('materialEfficiency').getCell(1, 1).getA1Notation();

		const curSheet = e.range.getSheet().getName();
		//const curSheet = 'Industry Calculator'; // TESTING
		if (curSheet === 'Industry Calculator') {
			const editedCell = e.range.getA1Notation();
			//const editedCell = componentNameCell; // TESTING
			if (
				editedCell === componentNameCell ||
				editedCell === optimizeForCell ||
				editedCell === materialEfficiencyCell ||
				editedCell === optimizeForCell ||
				reprocessingEfficiencyCells.includes(editedCell) ||
				oreToggles.includes(editedCell)
			) {
				setLoadingStatus();
				if (editedCell === componentNameCell) {
					loadComponentNeeds(e.range.getDisplayValue());
					//loadComponentNeeds('Dramiel'); // TESTING
				}
				if (calculateOresForComponent()) {
					unsetLoadingStatus();
				}
			}
		}
	} catch (error) {
		setErrorStatus();
		throw error;
	}
}

function getRangeA1Values(range) {
	const numRows = range.getNumRows();
	const numCols = range.getNumColumns();
	const rv = [];
	for (let i = 1; i <= numCols; i++) {
		for (let j = 1; j <= numRows; j++) {
			rv.push(range.getCell(j, i).getA1Notation());
		}
	}
	return rv;
}

function calculateOresForComponent() {
	getNeededMineralsFromSheet();
	const efficiencyTable = getDataArrayFromRangeByName('efficiencyTable');
	const engine = LinearOptimizationService.createEngine();
	initializeVariables(efficiencyTable, engine);
	constructConstraints(efficiencyTable, engine);
	configureObjectiveCoefficients(efficiencyTable, engine);
	engine.setMinimization();
	const solution = engine.solve();
	if (solution.isValid()) {
		const oreSolutionArr = getOreSolutionArray(efficiencyTable, solution);
		writeOreSolution(oreSolutionArr);
		return true;
	} else {
		const oreSolutionArr = getOreSolutionErrorArray(efficiencyTable);
		writeOreSolution(oreSolutionArr);
		setNotSolutionStatus();
		return false;
	}
}

function getDataArrayFromRangeByName(name) {
	const range = ActiveSpreadsheet.getRangeByName(name);
	const data = range.getValues();
	return data;
}

function initializeVariables(efficiencyTable, engine) {
	const oreEnabled = getOreToggles(efficiencyTable);
	//ore variables
	for (let i = 1; i < efficiencyTable.length; i++) {
		const oreName = efficiencyTable[i][0].toFormattedName();
		const upperBound = oreEnabled[oreName] ? 5 * 1000 * 1000 * 1000 * 1000 : 0;
		engine.addVariable(oreName, 0, upperBound, LinearOptimizationService.VariableType.CONTINUOUS);
	}
	//mineral variables
	for (let i = 1; i < efficiencyTable[i].length; i++) {
		const mineralName = efficiencyTable[0][i].toFormattedName();
		engine.addVariable(mineralName, 0, 5 * 1000 * 1000 * 1000 * 1000, LinearOptimizationService.VariableType.CONTINUOUS);
	}
}

function getOreToggles(efficiencyTable) {
	const oreToggleData = getDataArrayFromRangeByName('oreToggles');
	const rv = {};
	for (let row = 0; row < oreToggleData.length; row++) {
		for (let col = 0; col < oreToggleData[row].length; col++) {
			const oreName = efficiencyTable[row + 1][0].toFormattedName();
			rv[oreName] = oreToggleData[row][col];
		}
	}
	return rv;
}

function constructConstraints(efficiencyTable, engine) {
	for (let i = 1; i < efficiencyTable[i].length; i++) {
		const constraint = engine.addConstraint(minerals[efficiencyTable[0][i].toFormattedName()].needed, 5 * 1000 * 1000 * 1000 * 1000);
		for (let y = 1; y < efficiencyTable.length; y++) {
			const oreName = efficiencyTable[y][0].toFormattedName();
			constraint.setCoefficient(oreName, efficiencyTable[y][i]);
		}
	}
}

function configureObjectiveCoefficients(efficiencyTable, engine) {
	const optimizeFor = ActiveSpreadsheet.getRangeByName('optimizeFor').getCell(1, 1).getDisplayValue();

	switch (optimizeFor) {
		case 'Ore M3 Efficiency':
			configureOreM3ObjectiveCoefficients(efficiencyTable, engine);
			break;
		case 'Ore Cost Efficiency':
			configureOreCostObjectiveCoefficients(efficiencyTable, engine);
			break;
	}
}

function configureOreM3ObjectiveCoefficients(efficiencyTable, engine) {
	const volumeTable = getDataArrayFromRangeByName('volumeTable');
	for (let i = 1; i < volumeTable.length; i++) {
		const oreName = efficiencyTable[i][0].toFormattedName();
		engine.setObjectiveCoefficient(oreName, volumeTable[i][0]);
	}
}

function configureOreCostObjectiveCoefficients(efficiencyTable, engine) {
	const oreIskPerUnitTable = getDataArrayFromRangeByName('oreIskPerUnit');
	for (let i = 1; i < oreIskPerUnitTable.length; i++) {
		const oreName = efficiencyTable[i][0].toFormattedName();
		engine.setObjectiveCoefficient(oreName, 100 * oreIskPerUnitTable[i][0]);
	}
}

function getNeededMineralsFromSheet() {
	const range = ActiveSpreadsheet.getRangeByName('neededMinerals');
	const data = range.getValues();
	for (let i = 0; i < data[0].length; i++) {
		minerals[data[0][i].toFormattedName()].needed = data[1][i];
	}
}

function getOreSolutionArray(efficiencyTable, solution) {
	const rv = [];
	for (let i = 1; i < efficiencyTable.length; i++) {
		const oreName = efficiencyTable[i][0].toFormattedName();
		const val = [Math.ceil(solution.getVariableValue(oreName)) * 100];
		rv.push(val);
	}
	return rv;
}

function getOreSolutionErrorArray(efficiencyTable) {
	const rv = [];
	for (let i = 1; i < efficiencyTable.length; i++) {
		rv.push(['0']);
	}
	return rv;
}

function writeOreSolution(oreSolutionArray) {
	ActiveSpreadsheet.getRangeByName('oreSolution').setValues(oreSolutionArray);
}

function loadComponentNeeds(componentName) {
	const data = getComponentData(componentName);
	writeComponentData(data);
}

function getComponentData(componentName) {
	const dataRange = ActiveSpreadsheet.getRangeByName('componentNames');
	const values = dataRange.getValues();
	const col = findComponentColumn(values, componentName);
	const componentValuesSheet = ActiveSpreadsheet.getSheetByName('Component Values');
	const componentData = componentValuesSheet.getRangeList([`${col}4:${col}11`, `${col}12:${col}49`, `${col}50:${col}50`]).getRanges();
	const rv = {
		minerals: getComponentMineralArray(componentData[0]),
		planetaryResources: getPlanetaryResourceArray(componentData[1]),
		manufactureCost: getManufactureCostArray(componentData[2]),
	};
	return rv;
}

function findComponentColumn(values, componentName) {
	for (let i = 0; i < values.length; i++) {
		for (let j = 0; j < values[i].length; j++) {
			if (values[i][j] == componentName) {
				return columnToLetter(j + 2);
			}
		}
	}
}

function columnToLetter(column) {
	let temp,
		letter = '';
	while (column > 0) {
		temp = (column - 1) % 26;
		letter = String.fromCharCode(temp + 65) + letter;
		column = (column - temp - 1) / 26;
	}
	return letter;
}

function getComponentMineralArray(mineralRange) {
	const data = mineralRange.getValues();
	const rv = [];
	const rowArr = [];
	for (let i = 0; i < data.length; i++) {
		for (let j = 0; j < data[i].length; j++) {
			rowArr.push(data[i][j]);
		}
	}
	rv.push(rowArr);
	return rv;
}

function getPlanetaryResourceArray(planetaryResourceRange) {
	return planetaryResourceRange.getValues();
}

function getManufactureCostArray(manufactureCostRange) {
	return manufactureCostRange.getValues();
}

function writeComponentData(componentData) {
	ActiveSpreadsheet.getRangeByName('mineralBaseNeeds').setValues(componentData.minerals);
	ActiveSpreadsheet.getRangeByName('planetaryMaterialBaseNeeds').setValues(componentData.planetaryResources);
	ActiveSpreadsheet.getRangeByName('manufactureCost').getCell(1, 1).setValues(componentData.manufactureCost);
}

function setLoadingStatus() {
	const statusBar = ActiveSpreadsheet.getRangeByName('statusBar');
	statusBar.getCell(1, 1).setValue('Calculating please wait . . .');
	statusBar.setBackground('#E69138');
}

function unsetLoadingStatus() {
	const statusBar = ActiveSpreadsheet.getRangeByName('statusBar');
	statusBar.getCell(1, 1).setValue('');
	statusBar.setBackground('#3D85C6');
}

function setErrorStatus() {
	const statusBar = ActiveSpreadsheet.getRangeByName('statusBar');
	statusBar.getCell(1, 1).setValue('ERROR: There was a problem processing your selection.');
	statusBar.setBackground('#CC0000');
}

function setNotSolutionStatus() {
	const statusBar = ActiveSpreadsheet.getRangeByName('statusBar');
	statusBar.getCell(1, 1).setValue('ERROR: A solution was not possible. Try enabling more ores.');
	statusBar.setBackground('#CC0000');
}
