# VB6.Activex
Componentes e Bibliotecas ActiveX


# Criação De Um Controle ActiveX Que É Uma Fonte De Dados

Começando com VB6, você pode criar um controle ActiveX que funciona como uma fonte de dados. Um controle de fonte de dados fornece campos de um Recordset para o qual outros controles podem se ligar. Exemplos de controles de fonte de dados que vêm out-of-the-box com VB seria o controle de dados tradicionais e do ActiveX Data Control (discutido no Capítulo 8 ).
O mínimo que você precisa fazer para implementar um controle como fonte de dados é: 
1.	Definir o UserControl 's DataSourceBehavior propriedade.
2.	Programar o UserControl 's GetDataMember evento para retornar uma referência a um objeto Recordset. Este evento é acionado sempre que um consumidor de dados (geralmente um controle acoplado) tem o seu DataSource propriedade definida para apontar para o controle da fonte de dados.

Estas duas etapas são suficientes se o comportamento do seu controle de dados será muito bem restrito, isto é, os programadores que utilizam a fonte de dados não pode determinar o tipo de conexão de dados, nem os dados que a fonte de dados expõe. Neste cenário restrito, o GetDataMember procedimento de evento irá se conectar a um conjunto codificado de registros em um banco de dados codificados usando um hard-coded motorista de dados. No entanto, você pode querer dar aos programadores de sua escolha da fonte de dados mais sobre como o controle se conecta a dados. Nesse caso, você vai querer dar aos programadores mais das características que o padrão Microsoft controles fonte de dados fornecer, a saber:
  *	Propriedades que permitem que o programador para especificar strings conectar eo texto de consultas que recuperam dados para criar conjuntos de registros específicos. O GetDataMember procedimento de evento, então, dinamicamente ler essas propriedades para inicializar e retornar o Recordset.
  *	A propriedade Recordset para que os programadores podem manipular diretamente Recordset seu controle fonte de dados em seu próprio código.

## O evento GetDataMember
O GetDataMember evento ocorre quando o DataSource propriedade de um consumidor de dados que depende do controle atual está definido. O objetivo deste evento é para retornar uma referência a um objeto Recordset válida através do seu segundo parâmetro. Este Recordset torna-se então disponível para o consumidor de dados que causou o evento ao fogo, em primeiro lugar. Na Listagem 13,17, o código em um GetDataMemberprocedimento de evento sempre retorna uma referência ao conjunto de registros do Empregados de mesa na Nwind banco de dados Access.

### LISTAGEM 13,17 
o procedimento de evento GetDataMember
```
Private Sub UserControl_GetDataMember(DataMember As String, Data As Object)
On Error GoTo GetDataMemberError
    ' rs and cn are Private variables of the UserControl
    ' se esta é a primeira vez, então eles não foram definidas ainda
    If RS Is Nothing Or cn Is Nothing Then
        ' Create a Connection object and establish a connection.
        Set cn = New ADODB.Connection       
        cn.ConnectionString = _
            "Provider=Microsoft.Jet.OLEDB.3.51;” & _
            “Data Source=c:\northwind.mdb"       
        cn.Open
        ' Create a RecordSet object.
        Set RS = New ADODB.Recordset        
        RS.Open "employees", cn, adOpenKeyset, adLockPessimistic
        RS.MoveFirst
    End If    
    Set Data = RS 
    Exit Sub    
GetDataMemberError:
    MsgBox "Error: " & CStr(Err.Number) & vbCrLf & vbCrLf & Err.Description
    Exit Sub
End Sub
```

Nota na listagem que você manipular dois UserControl privada variáveis: A variável que representa a conexão ADO e outro representando o Recordset ADO. O código de procedimento de evento só tem que configurá-los na primeira passagem através do procedimento. Uma vez que a conexão eo Recordset são inicializados, você pular o código de inicialização.
No final da rotina (mesmo antes da Exit Sub para se desviar de o manipulador de erro), o código atribui o conjunto de registros para o segundo parâmetro, dados. Isso na verdade retorna o Recordset para qualquer consumidor de dados acaba solicitado (tipicamente outro controle que acaba de ter seu DataSource propriedade definida para apontar para uma instância desse controle).

# Passos para criar um Controle da Fonte de Dados
Os passos seguintes irão implementar um controle da fonte de dados em pleno funcionamento:
## PASSO A PASSO
13,4 Criando um Controle da Fonte de Dados
1.	Criar um novo projeto de controle ActiveX.
2.	Definir uma referência no projeto para a biblioteca de dados apropriado por meio do Projeto, caixa de diálogo Referências menu.
3.	Definir o UserControl 's DataSourceBehavior propriedade para 1-vbDataSource.
4.	Criar procedimentos de propriedade para propriedades personalizadas que os programadores vão usar para manipular a conexão do controle da fonte de dados para dados. Normalmente, você vai implementar Cordas propriedades como ConnectString (seqüência de conexão para inicializar um objeto Connection) e RecordSource (string para realizar a consulta para inicializar os dados no conjunto de registros). Criar variáveis privadas para manter os valores de cada uma das propriedades. Criar constantes privada para realizar os seus valores padrão iniciais. O programa InitProperties, ReadProperties, e WriteProperties procedimentos de evento para persistem essas propriedades.
5.	Se você deseja expor Recordset do controle para que outros programadores manipular, então você deve criar um nome de propriedade personalizada, RecordSet. Seu tipo será o tipo apropriado de registros que você planeja programa para seu controle. Você pode optar por fazê-lo somente leitura, caso em que você só precisa dar-lhe uma Property Get procedimento. Declare uma variável de objeto privada para realizar o seu valor usando WithEvents (isso expõe os procedimentos de evento para outros programadores).
6.	Declare uma variável privada do tipo apropriado de conexão que você planeja programa para seu controle. Ele não vai corresponder a uma propriedade personalizada, mas é necessário para hospedar os Recordset.
7.	Código do InitProperties, ReadProperties, e WriteProperties eventos para gerir adequadamente e persistir os valores das propriedades criado nas etapas anteriores.
8.	Programar o UserControl 's GetDataMember procedimento de evento para inicializar um conjunto de registros e devolvê-lo no segundo parâmetro. Você vai obter o conjunto de registros ou a partir de informações contidas em personalizado e privado variáveis ou de hard-coded informações no GetDataMember procedimento de evento em si (consulte a seção anterior para um exemplo). Você deve executar algumas errortrapping para garantir que você realmente tem uma conexão válida.
9.	Colocar o código no UserControl 's Terminate evento que graciosamente fechar a conexão de dados.
10.	Se você quiser permitir que os usuários naveguem dados por manipular diretamente o seu UserControl , em seguida, colocar a interface do usuário apropriado em seu UserControl junto com o código para navegar a variável Recordset.
11.	Seu novo controle ActiveX deve agora estar pronto para testar como um DataSource: Adicionar um projeto EXE padrão para o grupo de projecto.Agora, certificando-se que você fechou o designer para o UserControl, adicione uma instância do seu novo controle para formar o EXE norma.
12.	Manipular as propriedades necessárias personalizado (como ConnectString ou RecordSource) que você pode ter colocado no seu controle personalizado.
13.	Coloque um ou mais controles bindable no projeto de teste e definir suas DataSource propriedade para apontar para a instância do seu controle de fonte de dados. Definiu seu DataField propriedades para apontar para campos do Recordset expostos.
