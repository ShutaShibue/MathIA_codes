const filename = 'data.xlsx'
const area = "B3:H125"
import { write } from './module'
//--------settings

async function sevenpair_1p(Pairs: number) {
    let Round = 0,
        Prob = 0;
    
    while(Pairs < 7){
        do {
            Prob = (39 - 6 * Pairs) / (123 - Round);
            Round++;
        } while (Prob <= Math.random())
        
        Pairs++
    }    
    return Round
}

async function sevenpair_4p(Pairs: number) {
    let Round = 0,
        Prob = 0,
        Remain_tile = 123; 
    
    while (Pairs < 7) {
        do {
            if (Remain_tile < 0) return false;
            Prob = Pairs === 6
                ? Prob = 1 - (1 - 2 / Remain_tile) ** 4
                : 3 * (2 * (6 - Pairs) + 1) / (Remain_tile);
            
            Remain_tile -= 4; 
            Round++; 
        } while (Prob < Math.random())

        Pairs++
    }    
    return Round
}
//-----
let data = new Array(7)

async function main() {
    /*
    for (let n = 0; n < 7; n++) {
        data[n] = new Array(123).fill(0);
        for (let m = 0; m < 10000000; m++) {
            const Round = await sevenpair_1p(n)
            data[n][Round-1]++
        }
    }
    await write(data, area, '1P', filename)
*/

    for (let n = 0; n < 7; n++) {
        data[n] = new Array(19).fill(0);
        for (let m = 0; m < 10000000; m++) {
            const k = await sevenpair_4p(n)
            if (k!==false) data[n][k-1]++
        }
    }
    await write(data, area, '4P', filename)
}

main()
