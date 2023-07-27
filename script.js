let table = document.getElementById("table");
let firstRow = document.getElementById("first-row");
let currentCellEle = document.getElementById("current-cell");
let tBody = document.getElementById("tbody");

let boldEle = document.getElementById("bold");
let italicEle = document.getElementById("italic");
let colorTextEle = document.getElementById("text-color");
let cellBgEle = document.getElementById("cell-bg");
let alignLeftEle = document.getElementById("align-left");
let alignCenterEle = document.getElementById("align-center");
let alignRightEle = document.getElementById("align-right");

let displayTextColor = document.getElementById("display-text-color");
let displayBgColor = document.getElementById("display-cell-bg-color");

let cutEle = document.getElementById("cut");
let copyEle = document.getElementById("copy");
let pasteEle = document.getElementById("paste");

let minusEle = document.getElementById("minus");
let plusEle = document.getElementById("plus");
let fontSizeTextEle = document.getElementById("font-size-text");

let fontFamilyEle = document.getElementById("font-family");

let downloadEle = document.getElementById("download");
let uploadEle = document.getElementById("upload");
let uploadInputEle = document.getElementById("upload-input");

let displaySheetNoEle = document.getElementById("display-sheet-no");
let addSheetEle = document.getElementById("add-new-sheet");
let sheetBtnCont = document.getElementById("sheet-btn-cont");


let charA = 64;
let cutCell = {};
let currentSheetNo = 1;
let totalSheet = 1;


/* Virtual Memory*/

let matrix = new Array(100);

for(let row = 0; row < 100; row++){
	matrix[row] = new Array(26);
	for(let col = 0; col < 26; col++){
		matrix[row][col] = {};
	}
}

/* updating matrix */

function updateMatrix(currentCell){
	let tempObj = {
		style: currentCell.style.cssText,
		text: currentCell.innerText,
		id : currentCell.id
	}

	let j = currentCell.id[0].charCodeAt(0) - 65;
	let i = currentCell.id.substring(1) - 1;

	matrix[i][j] = tempObj;
}

/* download */

downloadEle.addEventListener("click", function(){
	let matrixStr= JSON.stringify(matrix);

	let blob = new Blob([matrixStr], {type: 'application/json'});

	let link = document.createElement("a");
	link.href = URL.createObjectURL(blob);
	link.download = "excelDownload.json";

	document.body.appendChild(link);
	link.click();
	document.body.removeChild(link);
})


/* upload */

uploadEle.addEventListener("click", function(){

	if(uploadInputEle.style.display === "none"){

		uploadInputEle.style.display = "block";

		uploadInputEle.addEventListener("change",function(event){
			let file = event.target.files[0];

			if(file){
				let reader = new FileReader();
				reader.readAsText(file);
				reader.onload = function(e){
					let fileContent = e.target.result;

					try{
						matrix = JSON.parse(fileContent);

						matrix.forEach((row) => {
							row.forEach((cell) => {
								if(cell.id){
									let currentCell = document.getElementById(`${cell.id}`);
									currentCell.innerText = cell.text;
									currentCell.style.cssText = cell.style;
								}
							})
						})
					}catch(err){
						console.log(err)
					}
				}
			}

			uploadInputEle.style.display = "none";
		})
	}else{
		uploadInputEle.style.display = "none";
	}
	
})


/* on clicking add button */

addSheetEle.addEventListener("click", function(){
	totalSheet = totalSheet + 1;
	currentSheetNo = totalSheet;

	let buttonEle = document.createElement("button");
	buttonEle.innerText = `Sheet ${currentSheetNo}`
	buttonEle.classList.add("sheet-btn");
	buttonEle.setAttribute("id", `sheet-${currentSheetNo}`);
	buttonEle.setAttribute("onclick", 'viewSheet(event)');

	sheetBtnCont.appendChild(buttonEle);
	displaySheetNoEle.innerText = `Current Sheet No : ${currentSheetNo}`;

	if(localStorage.getItem("arrMatrix")){
		
		var oldMatrix = localStorage.getItem("arrMatrix");
		var temp = [...JSON.parse(oldMatrix)]
		console.log(temp);
		var newMatrix = [...JSON.parse(oldMatrix), matrix];

		localStorage.setItem("arrMatrix", JSON.stringify(newMatrix));
	}else{
		
		let tempMatrixArr = [matrix];
		console.log("tempMatrixArr",tempMatrixArr);
		localStorage.setItem("arrMatrix", JSON.stringify(tempMatrixArr));
	}

	/* cleaning virtual memory-  matrix */

	for(let row = 0; row < 100; row++){
		matrix[row] = new Array(26);
		for(let col = 0; col < 26; col++){
			matrix[row][col]  = {};
		}
	}

	// clearing table

	tBody.innerHTML = "";
	console.log("test");

	addTableToSheet();


})


/* view Sheet Button */

function viewSheet(event){
	let id = event.target.id.split("-")[1];
	console.log("id clicked : ", id)

	if(Number(id) === totalSheet){
		currentSheetNo = Number(id);
		displaySheetNoEle.innerText = `Current Sheet No : ${currentSheetNo}`;
		alert("please save your final sheet")
	}else{

		var matrixArr = JSON.parse(localStorage.getItem("arrMatrix"));

		console.log("matrixArr",matrixArr);
		matrix = matrixArr[Number(id) - 1];
		console.log("matrix",matrix);

		tBody.innerHTML = "";

		addTableToSheet()

		matrix.forEach(row => {
			row.forEach(cell => {
				if(cell.id){
					let currentCell = document.getElementById(`${cell.id}`);
					currentCell.innerText = cell.text;
					currentCell.style.cssText = cell.style;
				}
			}) 
		})

		currentSheetNo = Number(id);
		displaySheetNoEle.innerText = `Current Sheet No : ${currentSheetNo}`;
	}
}


/* Bold */

boldEle.addEventListener("click", function(){
	let currentCell = document.getElementById(`${currentCellEle.innerText}`);
	if(currentCell.style.fontWeight === "normal"){
		currentCell.style.fontWeight = "bold";
	}else{
		currentCell.style.fontWeight = "normal";
	}

	updateMatrix(currentCell)
})

/* Italic */

italicEle.addEventListener("click", function(){
	let currentCell = document.getElementById(`${currentCellEle.innerText}`);
	if(currentCell.style.fontStyle === "Italic"){
		currentCell.style.fontStyle = "normal";
	} else{
		currentCell.style.fontWeight = "Italic";
	}

	updateMatrix(currentCell)
})

/* Text color */

colorTextEle.addEventListener("click", function(){

	let currentCell = document.getElementById(`${currentCellEle.innerText}`);

	if(displayBgColor.style.display === "block"){
		displayBgColor.style.display = "none";
	}


	if(displayTextColor.style.display === "block"){
		displayTextColor.style.display = "none";
	}else{
		displayTextColor.style.display = "block";
		displayTextColor.addEventListener("input", function(){
			
			currentCell.style.color = displayTextColor.value;
		})

		displayTextColor.addEventListener("blur",function(){
			displayTextColor.style.display = "none";
		})
	}

	updateMatrix(currentCell)

	
})

/* Cell bg */

cellBgEle.addEventListener("click", function(){

	let currentCell = document.getElementById(`${currentCellEle.innerText}`);

	if(displayTextColor.style.display = "block"){
		displayTextColor.style.display = "none";
	}

	if(displayBgColor.style.display === "block"){
		displayBgColor.style.display = "none";
	}else{
		displayBgColor.style.display = "block";
		displayBgColor.addEventListener("input", function(){
			
			currentCell.style.backgroundColor = displayBgColor.value;
		})

		displayBgColor.addEventListener("blur", function(){
			displayBgColor.style.display = "none";
		})
	}

	updateMatrix(currentCell)

	
})

/* align-left */

alignLeftEle.addEventListener("click", function(){
	let currentCell = document.getElementById(`${currentCellEle.innerText}`);
	currentCell.style.textAlign = "left";

	updateMatrix(currentCell);
})

/* align-center */
alignCenterEle.addEventListener("click", function(){
	let currentCell = document.getElementById(`${currentCellEle.innerText}`);
	currentCell.style.textAlign = "center";

	updateMatrix(currentCell);

})

/* align-right */

alignRightEle.addEventListener("click", function(){
	let currentCell = document.getElementById(`${currentCellEle.innerText}`);
	currentCell.style.textAlign = "right";

	updateMatrix(currentCell);
})

/* cut */

cutEle.addEventListener("click", function(){
	let currentCell = document.getElementById(`${currentCellEle.innerText}`);

	cutCell = {
		style: currentCell.style.cssText,
		text: currentCell.innerText
	}

	currentCell.innerText = "";
	currentCell.style.cssText = "";

	updateMatrix(currentCell);

})

/* copy */

copyEle.addEventListener("click", function(){
	let currentCell = document.getElementById(`${currentCellEle.innerText}`);

	cutCell = {
		style: currentCell.style.cssText,
		text: currentCell.innerText
	}


})

/* paste */

pasteEle.addEventListener("click", function(){
	let currentCell = document.getElementById(`${currentCellEle.innerText}`);

	currentCell.innerText = cutCell.text;
	currentCell.style.cssText = cutCell.style;

	updateMatrix(currentCell)
})

/* minus font size */

minusEle.addEventListener("click", function(){
	let curFontSize = Number(fontSizeTextEle.innerText);
	curFontSize = curFontSize - 1;

	fontSizeTextEle.innerText = curFontSize;

	let currentCell = document.getElementById(`${currentCellEle.innerText}`);
	currentCell.style.fontSize = `${fontSizeTextEle.innerText}px`;

	updateMatrix(currentCell)
})

/* plus font size */

plusEle.addEventListener("click", function(){
	let curFontSize = Number(fontSizeTextEle.innerText);
	curFontSize = curFontSize + 1;

	fontSizeTextEle.innerText = curFontSize;

	let currentCell = document.getElementById(`${currentCellEle.innerText}`);
	currentCell.style.fontSize = `${fontSizeTextEle.innerText}px`;

	updateMatrix(currentCell)
})

/* Font family */

fontFamilyEle.addEventListener("input", function(){
	let currentCell = document.getElementById(`${currentCellEle.innerText}`);
	currentCell.style.fontFamily = fontFamilyEle.value;

	updateMatrix(currentCell);
})


/*  on focus function */
function onFocusFn(event){
	let currentCell = event.target;
	currentCellEle.innerText = currentCell.id;
}


/* on Input function */

function onInputFn(event){
	let currentCell = event.target;
	updateMatrix(currentCell);
}


/* Adding table to sheet */

function addTableToSheet(){
	console.log("inside func")
	for(let i = 1; i <= 26; i++){
		let th = document.createElement("th");
		th.innerText = String.fromCharCode(charA + i);
		th.classList.add("th");

		firstRow.appendChild(th);
	}

	for(let row = 1; row <= 100; row++){
		let newRow = document.createElement("tr");
		let th = document.createElement("th");
		th.classList.add("th");
		th.innerText = row;

		newRow.appendChild(th);

		for(let col = 1; col <= 26; col++){
			let newCell = document.createElement("td");
			newCell.classList.add("td");
			newCell.setAttribute("contenteditable", 'true');
			newCell.id = `${String.fromCharCode(charA + col)}${row}`;

			newCell.addEventListener("input", (event) => onInputFn(event));
			newCell.addEventListener("focus", (event) => onFocusFn(event));

			newRow.appendChild(newCell);

		}

		tBody.appendChild(newRow);
	}
}


addTableToSheet();