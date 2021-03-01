import java.io.File;

   import javax.swing.JOptionPane;
   import java.io.FileInputStream;
   import java.io.FileOutputStream;
   import java.util.Iterator;
   import java.util.Map;
   import java.util.TreeMap;
   
   import org.apache.poi.hssf.usermodel.HSSFWorkbook;
   import org.apache.poi.ss.usermodel.Cell;
   import org.apache.poi.ss.usermodel.Row;
   import org.apache.poi.ss.usermodel.Sheet;
   import org.apache.poi.ss.usermodel.Workbook;

public class TriCLOUD {
//chemin du fichier excel qu'on veut traiter
	private static final String FILE_NAME= "C:\\TRIProject\\CLOUD.xls";
//fichier du tri par nom
	private static final String FILE_NAME_TRI1= "C:\\TRIProject\\CLOUDTriNom.xls";
//fichier du tri par prenom
	private static final String FILE_NAME_TRI2= "C:\\TRIProject\\CLOUDTriPrenom.xls";
//fichier du tri par note
	private static final String FILE_NAME_TRI3= "C:\\TRIProject\\CLOUDTriNote.xls";
	
	public static void main(String[] args) throws Exception {
		
		/*créer un rapport à partir de notre fichier CLOUD
		c'est à dire créer espace de travail similaire à une feuille excel*/
		FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
		Workbook originalWorkbook = new HSSWorkbook(excelFile);
		sheet originalSheet = originalWorkbook.getSheetAt(0); //sheet(0) est la 1 ere feuille excel du fichier CLOUD
		
		
		//choix du methode de tri apres affichage de la fenetre du dialogue 
		String s=JOptionPane.showInputDialog("VEUILLEZ CHOISIR UN TRI: \n\t1: Tri Par Prenom \n\t3: Tri par Note");
		int choix= Integer.parseInt(s);
		 
	/*ananlyse du choix de l'utilisateur 
	 * apres une presentation d'un menu des choix des possibilités de tri 
	 * par nom
	 * par prenom
	 * par note */
		switch(choix)
		case 1:
			// creer un SortedMap<String, Row> ou la clé est la valeur de la 1ere colonne 
		// ca va automatiquement trier les lignes (Treemap)
			Map<Srtring, Row>sortedRowsMap = ne TreeMap<>();
			Iterator<Row> rowIterator = originalSheet.rowIterator();
			while(rowIterator.hasNext()) {
				Row row = rowIterator.next();
				sortedRosMap.put(row.getCell(0).getStringCellValue(), row);	
			}
			//creer un nouveau workbook
			Workbook sortedWorkbook ) new HSSFWorkbook();
			Sheet sortedSheet = sortedWorkbook.createSheet(originalSheet.getSheetName());
			// copier toutes les lignes triées dans le nouveau workbook 
			int rowIndex= 0;
			for(Row row : sortedRowsMap.values()) {
				Row newRow= sortedSheet.createRow(rowIndex);
				copyRowToRow(row, newRow);//la methode RowtoRow est définie ulterieurement 
				rowIndex++;
				
			}
			//ecrire le nouveau workbook dans un nouveau fichier excel 
			// si le fichier excel n'exsite pas il va etre crée 
			FileOutputStream out = new FileOutputStream(FILE_NAME_TRI1);
			sortedWorkbook.wite(out);
			break;
			case 2:
				// creer un SortedMap<string, Row> ou la clé est la valeur de la 2 eme colonne de la feuille excel
				 Map(String, Row)sortedRowsMap1 = new TreeMap<> ();
				 Iterator<Row> rowIterator1 = originalSheet.rowIterator();
				 while (rowIterator1.hasNext()) {
					 Row row =rowIterator1.next();
					 sortedRowsMap1.put(row.getCell(1).getStringCellValue(), row); //2e colonne est d'indice 1
				 }
			// nv workbook
				 workbook sortedWorkbook1 = new HSSFWorkbook();
				 Sheet sortedSheet = sortedWorkbook1.createSheet(originalSheet.getSheetName());
			//copie des lignes triées dans le nouveau workbook
				 int rowIndex1 = 0 ;
				 for(Row row : sortedRowsMap1.values()) {
					 Row newRow = sortedSheet1.createRow(rowIndex1);
					 copyRowToRow(row, newRow);
					 rowIndex1++;
				 }
			//Ecriture dans excel 
				 FileOutputStream out1 ) new FileOUtpurStream(FILE_NAME_TRI2);
				 sortedWorkbook1.write(out1);
				 break;
				 case 3:
					// creer un SortedMap<string, Row> ou la clé est la valeur de la 3eme colonne de la feuille excel
					// ca va automatiquement trier les lignes (Treemap)
					 Map(String, Row)sortedRowsMap11 = new TreeMap<> ();
					 Iterator<Row> rowIterator11 = originalSheet.rowIterator();
					 while (rowIterator11.hasNext()) {
						 Row row =rowIterator11.next();
						 sortedRowsMap1.put(row.getCell(2).getStringCellValue(), row); //3e colonne est d'indice 2
					 }
				// creation nv workbook
					 workbook sortedWorkbook11 = new HSSFWorkbook();
					 Sheet sortedSheet = sortedWorkbook11.createSheet(originalSheet.getSheetName());
				//copie des lignes triées dans le nouveau workbook
					 int rowIndex11 = 0 ;
					 for(Row row : sortedRowsMap11.values()) {
						 Row newRow = sortedSheet11.createRow(rowIndex1);
						 copyRowToRow(row, newRow);
						 rowIndex1++;
					 }
				//Ecriture dans fichier.xsl
					 FileOutputStream out11 ) new FileOUtpurStream(FILE_NAME_TRI3);
					 sortedWorkbook11.write(out11);
					 break;
					 
					 default:
						 //message d'errreur de la saisie
						 //choix inexistentbdans le menu proposé
						 JOptionPane.showMessageDialog(null, "veuillez entrer une des valeurs proposées","ERREUR", JOptionPane.ERROR_MESSAGE);
	}
					 
		
		\\ la méthode pour copier ligne par ligne nomée RowToRow
	private static void copyRowToRow(Row row, Row newRow) {
		Iterator<Cell>cellIterator = row.cellIterator();
		int cellIndex = 0;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			Cell newCell = newRow.createCell(cellIndex);
			newCell.setCellValue(cell.getStringCellValue());
			cellindex++cell;
				
		}
	}
}