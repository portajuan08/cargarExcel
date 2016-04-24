package CargarExcel;
import java.awt.BorderLayout;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Principal extends JFrame{
	private static final long serialVersionUID = 1L;
	
	JTextField txtCargarCajas;
	JFileChooser selectorCargarCajas;
	JButton seleccionarCargarCajas;
	List<RenglonExcel> datosCargarCajas;
	
	JTextField txtCargarIntervalos;
	JFileChooser selectorCargarIntervalos;
	JButton seleccionarCargarIntervalos;
	Map<Double,Double> datosCargarIntervalos;
	
	JButton intercalarDatos;
	
	public Principal(){
		super("Buscar Cajas de Creditos");
		
		JPanel panel =new JPanel();
        panel.setLayout(new FlowLayout());
        setSize(1200,300);
        setLocationRelativeTo(null);
        setVisible(true);
        panel.add(new JLabel("Archivo con datos de las cajas: "));
        txtCargarCajas=new JTextField(15);
        panel.add(txtCargarCajas);
        seleccionarCargarCajas=new JButton("Seleccionar Archivo");
  
        seleccionarCargarCajas.addActionListener(new seleccionarExcel(true));
        panel.add(seleccionarCargarCajas);
        
        panel.add(new JLabel("Archivo con datos de los intervalos: "));
        txtCargarIntervalos=new JTextField(15);
        panel.add(txtCargarIntervalos);
        seleccionarCargarIntervalos=new JButton("Seleccionar Archivo");
  
        seleccionarCargarIntervalos.addActionListener(new seleccionarExcel(false));
        panel.add(seleccionarCargarIntervalos);
		  
        intercalarDatos =new JButton("Obtener nro de cajas de Intervalos");
        
        intercalarDatos.addActionListener(new cargarResultados(this));
        panel.add(intercalarDatos);
		add(panel,BorderLayout.WEST);
		  
	}
	
	
	public void cargarArchivo(File fileName,boolean bEsCargarDatos) throws ParseException{
		List<List<XSSFCell>> cellDataList = new ArrayList<List<XSSFCell>>();
		try{
			FileInputStream fileInputStream = new FileInputStream( fileName);
			XSSFWorkbook excelCargarCajas = new XSSFWorkbook(fileInputStream);
			XSSFSheet hssfSheet = excelCargarCajas.getSheetAt(0);
			Iterator<Row> rowIterator = hssfSheet.rowIterator();
			while (rowIterator.hasNext()){
				XSSFRow hssfRow = (XSSFRow) rowIterator.next();
				Iterator<Cell> iterator = hssfRow.cellIterator();
				List<XSSFCell> cellTempList = new ArrayList<XSSFCell>();
				while (iterator.hasNext()){
					XSSFCell hssfCell = (XSSFCell) iterator.next();
					cellTempList.add(hssfCell);
				}
				cellDataList.add(cellTempList);
			}
		}catch (Exception e)
		{e.printStackTrace();}
		if (bEsCargarDatos)
			LeerCargarCajas(cellDataList);
		else
			LeerCargarIntervalos(cellDataList);
	}
	
	private void LeerCargarCajas(List<List<XSSFCell>> cellDataList){
		datosCargarCajas = new ArrayList<RenglonExcel>();
		for (int i = 0; i < cellDataList.size(); i++){
			List<?> cellTempList = (List<?>) cellDataList.get(i);
			
			XSSFCell hssfCell = (XSSFCell) cellTempList.get(0);
			Double nroCaja = null;
			if (hssfCell.getCellType() != Cell.CELL_TYPE_BLANK)
				nroCaja = hssfCell.getNumericCellValue();
			
			hssfCell = (XSSFCell) cellTempList.get(1);
			Double desde = null;
			if (hssfCell.getCellType() != Cell.CELL_TYPE_BLANK)
				desde = hssfCell.getNumericCellValue();
			
			hssfCell = (XSSFCell) cellTempList.get(2);
			Double hasta = null;
			if (hssfCell.getCellType() != Cell.CELL_TYPE_BLANK)
				hasta = hssfCell.getNumericCellValue();
			datosCargarCajas.add(new RenglonExcel(nroCaja, desde, hasta));
		}
		JOptionPane.showMessageDialog(this, "Ya se cargaron todas las cajas");
		
	}
	
	private void LeerCargarIntervalos(List<List<XSSFCell>> cellDataList){
		datosCargarIntervalos = new HashMap<Double,Double>();
		for (int i = 0; i < cellDataList.size(); i++){
			List<?> cellTempList = (List<?>) cellDataList.get(i);
			
			XSSFCell hssfCell = (XSSFCell) cellTempList.get(0);
			Double intervalo = null;
			if (hssfCell.getCellType() != Cell.CELL_TYPE_BLANK)
				intervalo = hssfCell.getNumericCellValue();
			
			hssfCell = (XSSFCell) cellTempList.get(1);
			Double nroCaja = null;
			if (hssfCell.getCellType() != Cell.CELL_TYPE_BLANK)
				nroCaja = hssfCell.getNumericCellValue();
			datosCargarIntervalos.put(intervalo, nroCaja);
		}
		JOptionPane.showMessageDialog(this, "Ya se cargaron todos los intervalos");
	}
	
	private class seleccionarExcel implements ActionListener{
		
		private boolean esCargarDatos;
		
		public seleccionarExcel(boolean bEscargarDatos) {
			esCargarDatos = bEscargarDatos;
		}
		
		@Override
		public void actionPerformed(ActionEvent e) {
			if (esCargarDatos)
				cargarCaja();
			else
				cargarIntervalos();
		}
		
		private void cargarCaja(){
			selectorCargarCajas=new JFileChooser();
			int op=selectorCargarCajas.showOpenDialog(Principal.this);
			if(op==JFileChooser.APPROVE_OPTION){
				try {
					txtCargarCajas.setText(selectorCargarCajas.getSelectedFile().getName());
					cargarArchivo(selectorCargarCajas.getSelectedFile(),esCargarDatos);
				} catch (Exception e1) {
					e1.printStackTrace();
				}
		    }
		}
		
		private void cargarIntervalos(){
			selectorCargarIntervalos=new JFileChooser();
			int op=selectorCargarIntervalos.showOpenDialog(Principal.this);
			if(op==JFileChooser.APPROVE_OPTION){
				try {
					txtCargarIntervalos.setText(selectorCargarIntervalos.getSelectedFile().getName());
					cargarArchivo(selectorCargarIntervalos.getSelectedFile(),esCargarDatos);
				} catch (Exception e1) {
					e1.printStackTrace();
				}
		    }
		}
	}
	
	private class cargarResultados implements ActionListener{
		
		private Principal principal;
		
		public cargarResultados(Principal principal) {
			this.principal = principal;
		}
		
		@Override
		public void actionPerformed(ActionEvent e) {
			if (datosCargarCajas == null){
				JOptionPane.showMessageDialog(principal, "Faltan cargar los datos de las cajas.");
				return;
			}
			if (datosCargarIntervalos == null){
				JOptionPane.showMessageDialog(principal, "Faltan cargar los intervalos.");
				return;
			}
			
			JFileChooser selectorResultados = new JFileChooser();
			selectorResultados.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
			int op=selectorResultados.showOpenDialog(Principal.this);
			if(op==JFileChooser.APPROVE_OPTION){
				for (Double intervalo : datosCargarIntervalos.keySet()){
					Double nroCaja = null;
					for (RenglonExcel oReng : datosCargarCajas){
						if (intervalo >= oReng.getDesde() && intervalo <= oReng.getHasta() ){
							nroCaja = oReng.getNroCaja();
							break;
						}
					}
					datosCargarIntervalos.put(intervalo, nroCaja);
				}
			
				HSSFWorkbook libro = new HSSFWorkbook();
				HSSFSheet hoja = libro.createSheet();
				int i = 0;
				for (Double intervalo : datosCargarIntervalos.keySet()){
					HSSFRow fila = hoja.createRow(i);
					HSSFCell celda1 = fila.createCell((short)0);
					HSSFRichTextString texto = new HSSFRichTextString(intervalo.toString());
					celda1.setCellValue(texto);
					HSSFCell celda2 = fila.createCell((short)1);
					if (datosCargarIntervalos.get(intervalo) != null)
						texto = new HSSFRichTextString(datosCargarIntervalos.get(intervalo).toString());
					else
						texto = new HSSFRichTextString("ERROR");		
					celda2.setCellValue(texto);
					i++;
				}
				try {
					FileOutputStream elFichero = new FileOutputStream(selectorResultados.getSelectedFile().getPath() + "/resultado.xlsx");
					libro.write(elFichero);
					elFichero.close();
					JOptionPane.showMessageDialog(principal, "Se creo un archivo resultado.xlsx con los resultados de los intervalos");
				} catch (Exception e1) {
					e1.printStackTrace();
				}
			}
		}
	}
	
	public static void main (String [ ] args) {
		new Principal();

	}	
}