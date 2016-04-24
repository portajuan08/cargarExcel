package CargarExcel;

public class RenglonExcel{
	Double nroCaja;
	Double desde;
	Double hasta;
	
	public RenglonExcel(Double dNroCaja, Double dDesde,Double dHasta){
		nroCaja = dNroCaja;
		desde = dDesde;
		hasta = dHasta;
	}
	
	public Double getNroCaja(){
		return nroCaja;
	}
	
	public Double getDesde(){
		return desde;
	}

	public Double getHasta(){
		return hasta;
	}

	public String toString(){
		return "nro caja => " + nroCaja + " desde => " + desde + " hasta => " + hasta;
	}
}
