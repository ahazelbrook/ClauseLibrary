Office.onReady(() => { 
    document.getElementById("clause-list").innerHTML = ` 
    <div> 
        <h3>Indemnification</h3> 
        <p>Standard indemnification clause text.</p> 
        <button onclick="insertClause('Standard indemnification clause text.')">Insert</button> 
    </div> 
    <div> 
        <h3>Initial Term and Renewals</h3> 
        <p>Initial term and renewal selections.</p> 
        <button onclick="insertClause('Initial term and renewal selections.')">Insert</button> 
    </div> 
    `;
 });
 
 function insertClause(text) {
    Word.run(async (context) => {
        const doc = context.document; 
        doc.body.insertText(text, Word.InsertLocation.end); 
        await context.sync();
    }); 
}