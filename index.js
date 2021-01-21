const keys = require ('./credentials.json');
const {google} = require('googleapis');
const { Console } = require('console');

const client = new google.auth.JWT(
  keys.client_email,
  null,
  keys.private_key,
  ['https://www.googleapis.com/auth/spreadsheets']  
);

client.authorize(function(err,tokens){
  if(err){
    console.log(err);
    return;
  }
  else{
    console.log('Connected!');
    Main(client);
  }
})

/**
 * Prints the names and majors of students in a sample spreadsheet:
 * @see https://docs.google.com/spreadsheets/d/11PP-1OmGt1Qta-wKa1XBcOzzb6Paia855yeDxDQmW1I/edit?usp=sharing
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */

function Main(auth) {
  const sheets = google.sheets({version: 'v4', auth});
  sheets.spreadsheets.values.get({
    spreadsheetId: '11PP-1OmGt1Qta-wKa1XBcOzzb6Paia855yeDxDQmW1I',
    range: '!A4:H',
  }, (err, res) => {
    if (err) return console.log('The API returned an error: ' + err);
    const rows = res.data.values;   
    //Set "0" in the seventh column.  
    let newDataArray = rows.map((row) => {    
        row[7] = 0;
      return row;
    });

    //#region Update
    //Does an update on the spreadsheet.
    const updateOptions = {
      spreadsheetId: '11PP-1OmGt1Qta-wKa1XBcOzzb6Paia855yeDxDQmW1I',
      range: '!A4:H',
      valueInputOption: 'USER_ENTERED',
      resource: {values: newDataArray}
    }
    res =  sheets.spreadsheets.values.update(updateOptions);
    //#endregion

    //#region Average Score
    //Calculate the average score and put the current situation.
    console.log("Média:");
    rows.map((row)=>{
      //Check the absence.
      if(row[2]>(60*0.25)){ 
        row[6] = 'Reprovado por Falta';
      }
      else{
        var m = ((parseFloat(row[3])+parseFloat(row[4])+parseFloat(row[5]))/3).toFixed(2);
        //Round the score.
        if((50-parseFloat(m)<=1.5)&&(50-parseFloat(m)>0)){
          m = parseFloat(m) + (50-parseFloat(m));
        }
        else if((70-parseFloat(m)<=1.5)&&(70-parseFloat(m)>0)){
          m = parseFloat(m) + (70-parseFloat(m));
        }
        
        console.log(`${row[1]}: `+m);
        if(m<50){
          row[6] = 'Reprovado por Nota';
        }
        else if((m>=50)&&(m<70)){
          row[6] = 'Exame Final';
          var naf =  (100 - parseFloat(m)).toFixed(2);
          row[7] = naf;
        }
        else{
          row[6] = 'Aprovado';
        } 
      }
    });
    //#endregion

    if (rows.length) {    
      rows.map((row) => {
        // Print columns A and H, which correspond to indices 0 and 7.
        console.log('--------------------------------------------------------------------------------------------------------------------------------------------------');
        console.log(`Matricula: ${row[0]}|Aluno: ${row[1]}|Faltas: ${row[2]}|P1: ${row[3]}|P2: ${row[4]}|P3: ${row[5]}|Situação: ${row[6]}| Naf:${row[7]}`);
      });
    } else {
      console.log('No data found.');
    }
  });
}