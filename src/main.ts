declare const XLSX: any;

interface Cards {
    'Card ID': string,
    'Collector Number': number,
    Description: string,
    Image: string,
    Name: string,
    Set: string
}

async function update(): Promise<void> {
    try {
        // Fetch card list from xlsx file.
        const response: Response = await fetch('cards.xlsx');
        const arrayBuffer: ArrayBuffer = await response.arrayBuffer();
        const data: Uint8Array = new Uint8Array(arrayBuffer);

        const myWorkbook = XLSX.read(data, { type: 'array'});
        const mySheet = myWorkbook.Sheets[myWorkbook.SheetNames[0]];

        // Parse data to JSON.
        const jsonData: Array<Cards> = XLSX.utils.sheet_to_json(mySheet, { header: 2 });
        
        // Sort cards based on Collector Number.
        const sortOpt = document.getElementById('sort') as HTMLSelectElement;
        let sortedData: Array<Cards> = [];
        if (sortOpt) {
            if (sortOpt.options[sortOpt.selectedIndex].value === '1') {
                sortedData = jsonData.sort((a: Cards, b: Cards) => a['Collector Number'] - b['Collector Number']);
            } else {
                sortedData = jsonData.sort((a: Cards, b: Cards) => b['Collector Number'] - a['Collector Number']);
            }
        }
        
        // Generate each card's HTML content.
        const grid = document.getElementById('cards') as HTMLDivElement;
        if (grid) {
            grid.innerHTML = '';
            sortedData.forEach((card) => {
                const cell: HTMLDivElement = document.createElement('div');
                cell.setAttribute('class','card');
                cell.innerHTML = `
                    <img class="card-image" src=${card.Image} alt=${card.Name} />
                    <p class="card-num">${card.Set} | ${('00' + card["Collector Number"]).slice(-3)}</p>
                    <h2 class="card-name">${card.Name}</h2>
                    <p class="card-desc">${card.Description}</p>
                `
                grid.appendChild(cell);
            })
        }
    } catch (error) {
        console.error(error);
    }
}
