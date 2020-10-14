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

const planetaryMaterials = {};

const ActiveSpreadsheet = SpreadsheetApp.getActive();

// eslint-disable-next-line no-unused-vars
function onEdit(e) {
	try {
		SpreadsheetApp.flush();
		const componentNameCell = ActiveSpreadsheet.getRangeByName('selectedComponentName').getCell(1, 1).getA1Notation();
		const costCalculationCell = ActiveSpreadsheet.getRangeByName('costCalculationSetting').getCell(1, 1).getA1Notation();
		const costCalculationSettingValue = ActiveSpreadsheet.getRangeByName('costCalculationSetting').getCell(1, 1).getDisplayValue();
		const reprocessingEfficiencyCells = getRangeA1Values(ActiveSpreadsheet.getRangeByName('oreReprocessingEfficiencyValues'));
		const oreToggles = getRangeA1Values(ActiveSpreadsheet.getRangeByName('oreToggles'));
		const oreOnHandCells = getRangeA1Values(ActiveSpreadsheet.getRangeByName('oreOnHand'));
		const mineralsOnHandCells = getRangeA1Values(ActiveSpreadsheet.getRangeByName('mineralsOnHand'));

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
				editedCell === costCalculationCell ||
				reprocessingEfficiencyCells.includes(editedCell) ||
				(oreToggles.includes(editedCell) && costCalculationSettingValue == 'Remainder') ||
				(oreOnHandCells.includes(editedCell) && costCalculationSettingValue == 'Remainder') ||
				mineralsOnHandCells.includes(editedCell)
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
	getNeededMineralsBasedOnMode();
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

function getNeededMineralsBasedOnMode() {
	const data = getDataArrayFromRangeByName('neededMinerals');
	const costCalculationSettingValue = ActiveSpreadsheet.getRangeByName('costCalculationSetting').getCell(1, 1).getDisplayValue();

	switch (costCalculationSettingValue) {
		case 'Full Component':
			for (let i = 0; i < data[0].length; i++) {
				minerals[data[0][i].toFormattedName()].needed = data[2][i];
			}
			break;
		case 'Remainder':
			for (let i = 0; i < data[0].length; i++) {
				minerals[data[0][i].toFormattedName()].needed = data[1][i];
			}
			break;
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
	const borderLeft = ActiveSpreadsheet.getRangeByName('borderLeft');
	const borderRight = ActiveSpreadsheet.getRangeByName('borderRight');
	const borderBottom = ActiveSpreadsheet.getRangeByName('borderBottom');
	statusBar.getCell(1, 1).setValue('Calculating please wait . . .');
	statusBar.setBackground('#E69138');
	borderLeft.setBackground('#E69138');
	borderRight.setBackground('#E69138');
	borderBottom.setBackground('#E69138');
	SpreadsheetApp.flush();
}

function unsetLoadingStatus() {
	const statusBar = ActiveSpreadsheet.getRangeByName('statusBar');
	const borderLeft = ActiveSpreadsheet.getRangeByName('borderLeft');
	const borderRight = ActiveSpreadsheet.getRangeByName('borderRight');
	const borderBottom = ActiveSpreadsheet.getRangeByName('borderBottom');
	statusBar.getCell(1, 1).setValue('');
	statusBar.setBackground('#3D85C6');
	borderLeft.setBackground('#434343');
	borderRight.setBackground('#434343');
	borderBottom.setBackground('#434343');
	SpreadsheetApp.flush();
}

function setErrorStatus(message = 'ERROR: There was a problem processing your selection.') {
	const statusBar = ActiveSpreadsheet.getRangeByName('statusBar');
	const borderLeft = ActiveSpreadsheet.getRangeByName('borderLeft');
	const borderRight = ActiveSpreadsheet.getRangeByName('borderRight');
	const borderBottom = ActiveSpreadsheet.getRangeByName('borderBottom');
	statusBar.getCell(1, 1).setValue(message);
	statusBar.setBackground('#CC0000');
	borderLeft.setBackground('#CC0000');
	borderRight.setBackground('#CC0000');
	borderBottom.setBackground('#CC0000');
	SpreadsheetApp.flush();
}

function setNotSolutionStatus() {
	const statusBar = ActiveSpreadsheet.getRangeByName('statusBar');
	const borderLeft = ActiveSpreadsheet.getRangeByName('borderLeft');
	const borderRight = ActiveSpreadsheet.getRangeByName('borderRight');
	const borderBottom = ActiveSpreadsheet.getRangeByName('borderBottom');
	statusBar.getCell(1, 1).setValue('ERROR: A solution was not possible. Try enabling more ores.');
	statusBar.setBackground('#CC0000');
	borderLeft.setBackground('#CC0000');
	borderRight.setBackground('#CC0000');
	borderBottom.setBackground('#CC0000');
	SpreadsheetApp.flush();
}

// eslint-disable-next-line no-unused-vars
function craftComponent() {
	const enoughMinerals = checkHaveMineralsForComponent();
	const enoughPlanetaryMaterials = checkHavePlanertaryMaterialsForComponent();
	if (enoughMinerals && enoughPlanetaryMaterials) {
		setLoadingStatus();
		subtractComponentResources();
		writeComponentResources();
		calculateOresForComponent();
		unsetLoadingStatus();
	} else {
		setErrorStatus('Not enough resources to craft component. Do you need to reprocess ores?');
		SpreadsheetApp.flush();
		Utilities.sleep(5000);
		unsetLoadingStatus();
	}
}

function checkHaveMineralsForComponent() {
	const neededMinerals = getDataArrayFromRangeByName('neededMinerals');
	const onHandMinerals = getDataArrayFromRangeByName('mineralsOnHand');
	let rv = true;

	for (let i = 0; i < neededMinerals[0].length; i++) {
		minerals[neededMinerals[0][i].toFormattedName()].needed = neededMinerals[2][i];
		minerals[neededMinerals[0][i].toFormattedName()].current = onHandMinerals[i][0];
		if (neededMinerals[2][i] > onHandMinerals[i][0]) rv = false;
	}
	return rv;
}

function checkHavePlanertaryMaterialsForComponent() {
	const neededplanetaryMaterials = getDataArrayFromRangeByName('planetaryMaterialNeeds');
	const onHandPlanetaryMaterials = getDataArrayFromRangeByName('planetaryMaterialsOnHand');
	let rv = true;

	for (let i = 0; i < neededplanetaryMaterials.length; i++) {
		planetaryMaterials[neededplanetaryMaterials[i][0].toFormattedName()] = { needed: neededplanetaryMaterials[i][2], current: onHandPlanetaryMaterials[i][0] };
		if (neededplanetaryMaterials[i][2] > onHandPlanetaryMaterials[i][0]) rv = false;
	}
	return rv;
}

function subtractComponentResources() {
	checkHaveMineralsForComponent();
	checkHavePlanertaryMaterialsForComponent();
	for (const iter in minerals) {
		minerals[iter].current -= minerals[iter].needed;
	}
	for (const iter in planetaryMaterials) {
		planetaryMaterials[iter].current -= planetaryMaterials[iter].needed;
	}
}

function writeComponentResources() {
	const mineralArr = [];
	for (const iter in minerals) {
		mineralArr.push([minerals[iter].current]);
	}
	const planetaryMaterialArr = [];
	for (const iter in planetaryMaterials) {
		planetaryMaterialArr.push([planetaryMaterials[iter].current]);
	}
	ActiveSpreadsheet.getRangeByName('mineralsOnHand').setValues(mineralArr);
	ActiveSpreadsheet.getRangeByName('planetaryMaterialsOnHand').setValues(planetaryMaterialArr);
	SpreadsheetApp.flush();
}

// eslint-disable-next-line no-unused-vars
function reprocessOres() {
	setLoadingStatus();
	try {
		const onHandReprocessesArr = flattenDataArray(getDataArrayFromRangeByName('onHandReprocesses'));
		const onHandOresArr = flattenDataArray(getDataArrayFromRangeByName('oreOnHand'));
		const mineralsToAdd = calculateReprocessMinerals(onHandReprocessesArr);
		const onHandMinerals = flattenDataArray(getDataArrayFromRangeByName('mineralsOnHand'));
		const newMineralsOnHand = sumArrays(mineralsToAdd, onHandMinerals);
		ActiveSpreadsheet.getRangeByName('mineralsOnHand').setValues(flatArrayToSingleColumn(newMineralsOnHand));
		let newOresOnHand = [];
		for (let i = 0; i < onHandOresArr.length; i++) {
			newOresOnHand.push(onHandOresArr[i] - 100 * onHandReprocessesArr[i]);
		}
		ActiveSpreadsheet.getRangeByName('oreOnHand').setValues(flatArrayToSingleColumn(newOresOnHand));
	} catch (error) {
		setErrorStatus();
	}
	unsetLoadingStatus();
}

function calculateReprocessMinerals(onHandReprocessesArr) {
	const rv = [];
	for (const mineral in minerals) {
		const mineralReprocessData = flattenDataArray(getDataArrayFromRangeByName(`${mineral}ReprocessValues`));
		rv.push(Math.round(sumProductArrays(onHandReprocessesArr, mineralReprocessData)));
	}
	return rv;
}

function flattenDataArray(arr) {
	const rv = [];
	arr.map((row) => {
		row.map((cell) => {
			rv.push(cell);
		});
	});
	return rv;
}

function flatArrayToSingleColumn(arr) {
	const rv = [];
	arr.map((val) => {
		rv.push([val]);
	});
	return rv;
}

function sumProductArrays(left, right) {
	let sum = 0;
	for (let i = 0; i < left.length; i++) {
		sum += left[i] * right[i];
	}
	return sum;
}

function sumArrays(left, right) {
	let sum = [];
	for (let i = 0; i < left.length; i++) {
		sum.push(left[i] + right[i]);
	}
	return sum;
}
