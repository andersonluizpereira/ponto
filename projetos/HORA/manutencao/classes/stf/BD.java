package stf;
import java.sql.*;
import javax.swing.JOptionPane;
public class BD
{
    public static Connection connection = null;
    public static Statement statement = null;
    public static ResultSet resultSet = null; 
    public static final String DRIVER   = "sun.jdbc.odbc.JdbcOdbcDriver";
    public static final String URL      = "jdbc:odbc:Driver={Microsoft Access Driver (*.mdb)};DBQ=//10.50.112.88/c$/projetos/HORA/DB/CHS.mdb";

    /** 
     * m�todo que faz conex�o com o banco de dados
     * retorna true se houve sucesso, ou false em caso negativo
     */
    public static boolean getConnection()
    {
       try
       {
   	      Class.forName(DRIVER);
          connection = DriverManager.getConnection(URL , "", "");
          statement = connection.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);
          JOptionPane.showMessageDialog(null,"Conectou!");
          return true;
       }
       catch(ClassNotFoundException erro)
       {
       	  erro.printStackTrace();
          return false;
       }
       catch(SQLException erro)
       {
       	  erro.printStackTrace();
          return false;
       }
    }
    
    /**
     * Fecha ResultSet, Statement e Connection
     */
    public static void close()
    {
	   closeResultSet();
	   closeStatement();
	   closeConnection();	
	}
	
	private static void closeConnection()
	{
	   try
	   {
	      connection.close();
	      JOptionPane.showMessageDialog(null,"Desconectou");
	   }
	   catch(SQLException erro)
	   {
	      erro.printStackTrace();
	   } 
	}  

	private static void closeStatement()
	{
	   try
	   {
	      statement.close();
	   }
	   catch(Exception e)
	   {
          e.printStackTrace();
       }
	}

	private static void closeResultSet()
	{
	   try
	   {
	      resultSet.close();
	   }
	   catch(Exception e)
	   {
          e.printStackTrace();
	   }
	}
    
    /**
     * Carrega o resultSet com o resultado do script SQL
     */
    public static void setResultSet(String sql)
    {
	   try
	   {
		  resultSet = statement.executeQuery(sql);
	   }
       catch(SQLException erro)
       {
          erro.printStackTrace();
       } 
    }  

    /**
     * Executa um script SQL de atualiza��o
     * retorna um valor inteiro contendo a quantidade de linhas afetadas
     */
    public static int runSQL(String sql)
    {
	   int quant = 0; 
	   try
	   {
		  quant = statement.executeUpdate(sql);
	   }
	   catch(SQLException erro)
	   {
	   	  erro.printStackTrace();
	   } 
       return quant;
	}  
}
