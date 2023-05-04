<!DOCTYPE html>
<html>
<head>
	<title>Consulta de chamados</title>
	<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/css/bootstrap.min.css" integrity="sha384-xOolHFLEh07PJGoPkLv1IbcEPTNtaed2xpHsD9ESMhqIYd0nLMwNLD69Npy4HI+N" crossorigin="anonymous">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/selectize.js/0.15.2/css/selectize.bootstrap4.css" integrity="sha512-+dSrbTGc/L04jAcqxlcnZUPXjnyKy6fdgmWQdsRILqZgjMOW8YfwWCQsna7XGXJDSaDvKW0ill4FgRmWX2Ki8w==" crossorigin="anonymous" referrerpolicy="no-referrer" />
	<style>
		.titulo{
			background-color: blue;
			padding: 3rem 3rem 1rem 1rem;
			margin: -2rem -3rem 0 -3rem;
			text-align: center;
			color: white;
		}
		
		table {
		  font-family: arial, sans-serif;
		  border-collapse: collapse;
		  width: 100%;
		}

		td, th {
		  border: 1px solid #dddddd;
		  text-align: center;
		  padding: 8px;
		}

		tr:nth-child(even) {
		  background-color: #dddddd;
		}
	</style>
</head>
<body>
	<div class="titulo" style="font-size: 12pt;">
		<h1 style="font-size: 2rem;">Relatório de técnicos por periodo</h1>
	</div>
	<div class="container">

		<br>
		<form method="POST" action="">
			<div class="row">
				<div class="col-md-3">
					<label for="data_inicial">Data inicial:</label>
					<input class="form-control" type="date" name="data_inicial" id="data_inicial" required>
				</div>
				
				<div class="col-md-3">
					<label for="data_final">Data final:</label>
					<input class="form-control" type="date" name="data_final" id="data_final" required>
				</div>
				<div class="col-md-3">
					<label for="tecnico">Técnico:</label>
					<select class="form-control" name="tecnico" id="tecnico">
				
					<?php
					// Configuração de conexão com o banco de dados
					$servername = "localhost";
					$username = "glpi";
					$password = "P@ssword";
					$dbname = "glpi";
					
					// Conexão com o banco de dados
					$conn = mysqli_connect($servername, $username, $password, $dbname);
					// Verifica se a conexão foi bem sucedida
					if (!$conn) {
						die("Falha na conexão com o banco de dados: " . mysqli_connect_error());
					}
					
					// Consulta SQL para obter a lista de usuários ativos
					$sql = "SELECT id, firstname FROM glpi_users WHERE glpi_users.is_active = '1' and entities_id='3'";
					$resultado = mysqli_query($conn, $sql);
					echo mysqli_num_rows($resultado);
					// Verifica se a consulta retornou algum resultado
					if (mysqli_num_rows($resultado) > 0) {
						// Exibe as opções do select com os usuários ativos
						while ($dados = mysqli_fetch_assoc($resultado)) {
							echo '<option value="' . $dados['id'] . '">'. $dados['id']. ' - '. $dados['firstname'] . '</option>';
						}
					} else {
						echo '<option value="">Nenhum técnico encontrado</option>';
					}
					?>
					</select>
				</div>
				<div class="col-md-2">
					<label>&nbsp;</label>
					<input class="btn btn-primary form-control" type="submit" name="submit" value="Consultar">
				</div>
			</div>
		</form>
	
<br>
<?php
require_once './vendor/autoload.php';

if(isset($_GET['baixado'])){
	if(file_exists($_GET['baixado']))
	unlink($_GET['baixado']);

	echo "<script></script>";
	header ('Location:/busca_tec');

}
// Cria um objeto DateTime com a data atual
//$timestamp = strtotime('now');
setlocale(LC_TIME, 'Portuguese_Brazil.1252');
$dateString = strftime('%A, %d de %^B de %Y', strtotime('today'));
$dateString = utf8_encode(ucwords(strftime('%A')).', '.strftime('%d').' de '.ucwords(strftime('%B')).' de '.strftime('%Y'));



// Cria um novo objeto PHPWord
$phpWord = new \PhpOffice\PhpWord\PhpWord();
$section = $phpWord->addSection();
$footer = $section->addFooter();
$footer->addPreserveText('Pág {PAGE} de {NUMPAGES}', null, array('alignment' => 'right'));

//Configuração do estilo de fonte de Titulo, sub-titulo e paragrafos
$headerFontStyle = 'Cabecalho';
$subheaderFontStyle = 'Parag';
$phpWord->addFontStyle(
    $headerFontStyle,
    array('name' => 'Arial', 'size' => 14, 'bold' => true)
);
$phpWord->addFontStyle(
    $subheaderFontStyle,
    array('name' => 'Arial', 'size' => 12, 'bold' => true)
);
$paragraphStyle = array(
    'align' => 'center',
	'marginTop' => 50,
	'lineHeight' => 1.5
);
$imageStyle = array(
    'wrappingStyle' => 'behind',
	'width' => 70, 'height' => 80
);

//Configração de fonte para Tópicos
$fontStyle = new \PhpOffice\PhpWord\Style\Font();
$fontStyle->setBold(true);
$fontStyle->setName('Arial');
$fontStyle->setSize(12);

//Configração de fonte para conteudo
$contentfontStyle = array(
	'name' => 'Arial', 'size' => 12, 'lineHeight' => 1.5

);
// $contentfontStyle = new \PhpOffice\PhpWord\Style\Font();
// $contentfontStyle->setName('Arial');
// $contentfontStyle->setSize(12);

$imagePath = './logo.png';
$section->addImage($imagePath, $imageStyle);



//Adicionando Texto - Cabeçalho
$section->addText(
    'PREFEITURA MUNICIPAL DE ITABIRA',
    $headerFontStyle, $paragraphStyle
);
$section->addText(
    'SECRETARIA MUNICIPAL DE DESENVOLVIMENTO URBANO',
    $headerFontStyle, $paragraphStyle
);
$section->addText(
    'SUPERINTENDÊNCIA DE GEOPROCESSAMENTO
	',
    $headerFontStyle, $paragraphStyle
);
$section->addText(
    'DIRETORIA DE CADASTRO E INFORMAÇÃO',
    $headerFontStyle, $paragraphStyle
);





// Verifica se o formulário foi submetido
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Obtém a data inicial e final do formulário
    $data_inicial = $_POST["data_inicial"];
    $data_final = $_POST["data_final"];

    $datahora_inicial = $data_inicial .' 00:00:00';
    $datahora_final =  $data_final.' 23:59:59';

	$timestamp = strtotime($data_inicial);
	$mes_ano = strftime('%B_%Y', $timestamp);

	$data_inicial_formatada = date('d-m-Y', strtotime($data_inicial));
	$data_final_formatada = date('d-m-Y', strtotime($data_final));

	$section->addText(
		'Relatório de Atividades no período de '.$data_inicial_formatada.' até ' .$data_final_formatada,
		$subheaderFontStyle, $paragraphStyle
	);

     // ID do técnico
    $id_tecnico = $_POST["tecnico"];

    // Consulta SQL para obter o nome do técnico e a quantidade de chamados no período
    $sql = "SELECT glpi_users.firstname, glpi_users.realname, COUNT(*) AS total 
            FROM glpi_tickets
            INNER JOIN glpi_tickets_users ON glpi_tickets_users.tickets_id = glpi_tickets.id
            INNER JOIN glpi_users ON glpi_users.id = glpi_tickets_users.users_id
            WHERE glpi_tickets_users.type = 2 
            AND glpi_tickets_users.users_id = $id_tecnico 
            AND glpi_tickets.date BETWEEN '$datahora_inicial' AND '$datahora_final' 
            GROUP BY glpi_users.firstname, glpi_users.realname;
    ";

    $resultado = mysqli_query($conn, $sql);

    // Verifica se a consulta retornou algum resultado
    if (mysqli_num_rows($resultado) > 0) {
        // Obtém os dados da consulta
        $dados = mysqli_fetch_assoc($resultado);
		// Define o nome do arquivo DOCX que será gerado
		$filename = "Rel_Ativ_Mensais_". $dados['firstname']."_".$mes_ano.".docx";

		$servidor=$dados['firstname'] . " " . $dados['realname'];

        // Adiciona o nome do técnico e a quantidade de chamados no período ao documento
        $section->addTextBreak();
        $myTextElement = $section->addText("Servidora: " . $dados['firstname'] . " " . $dados['realname']);
        $myTextElement->setFontStyle($fontStyle);
        $section->addTextBreak();


        $myTextElement = $section->addText("Foram atendidos um total de " . $dados['total'] . " chamados. Segue abaixo um breve relato das principais atividades desenvolvidas ", $contentfontStyle);
        $section->addTextBreak();
        // Consulta SQL para obter a lista de títulos dos chamados no período
        $sql = "SELECT DISTINCT
        MIN(CASE 
          WHEN glpi_tickets.name LIKE '%Atividades Administrativas%' THEN 'Acompanhamento das atividades administrativas' 
          ELSE glpi_tickets.name 
        END) AS name,
        CASE 
          WHEN glpi_tickets.status IN (1, 2, 3) THEN 'EM ATENDIMENTO'
          WHEN glpi_tickets.status = 4 THEN 'PENDENTE'
          WHEN glpi_tickets.status IN (5, 6) THEN 'FECHADO'
          ELSE 'Desconhecido'
        END AS status,
        case
        when CONCAT(
		FLOOR(timestampdiff(SECOND, date, closedate)/86400), ' dias ',
		FLOOR(MOD(timestampdiff(SECOND, date, closedate), 3600 * 24) / 3600), ' horas ',
		FLOOR(MOD(timestampdiff(SECOND, date, closedate), 3600) / 60), ' minutos '
  ) is not null THEN CONCAT(
		FLOOR(timestampdiff(SECOND, date, closedate)/86400), ' dias ',
		FLOOR(MOD(timestampdiff(SECOND, date, closedate), 3600 * 24) / 3600), ' horas ',
		FLOOR(MOD(timestampdiff(SECOND, date, closedate), 3600) / 60), ' minutos '
  ) else '' end  AS diferenca
		FROM glpi_tickets
		INNER JOIN glpi_tickets_users ON glpi_tickets_users.tickets_id = glpi_tickets.id 
		WHERE glpi_tickets_users.type = 2 
  		AND glpi_tickets_users.users_id = $id_tecnico 
        AND glpi_tickets.date BETWEEN '$data_inicial' AND '$data_final'
		GROUP BY name, status";

        $resultado = mysqli_query($conn, $sql);

        // Adiciona a lista de títulos dos chamados no período ao documento
        // Adiciona os itens de lista numerados com o estilo definido

		if (mysqli_num_rows($resultado) > 0) {
			$texto = "";
			$count = 1;
			while ($dados = mysqli_fetch_assoc($resultado)) {
				$texto .= $count . " – " . $dados['name'] . " – ".$dados['status']." " .$dados['diferenca']. "\n";
				$count++;
			}
			$linhas = explode("\n", $texto);
			foreach ($linhas as $linha) {
				$MyLinhas = $section->addText($linha, $contentfontStyle);
				$MyLinhas->setFontStyle($contentfontStyle);
			}
		} else {
			$section->addText("Nenhum chamado encontrado.");
		}


		

		// Adiciona nome e data no final do documento para assinatura
		$section->addTextBreak();
		$section->addTextBreak();
		$section->addText($servidor, $subheaderFontStyle, $paragraphStyle);
		$section->addText('Itabira – '.$dateString, $subheaderFontStyle, $paragraphStyle);

		
		

		// Cria um objeto de gravação do PHPWord para salvar o documento em formato DOCX
		 $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
		 $objWriter->save($filename);

		 echo "<a id='BtBaixar' href='http://localhost\Busca_tec/".$filename."' class ='btn btn-primary' download>Download</a>";

		echo '<div hidden><form > <input type="text" value="'.$filename.'" name="baixado">
					  <input type="submit" id="baixado"> </form></div>';

		echo "<b>".$filename."</b>";
	};
};

// Fecha a conexão com o banco de dados
mysqli_close($conn);
?>
	</div>
</body>
<footer>
<script src="https://code.jquery.com/jquery-3.6.4.min.js" integrity="sha256-oP6HI9z1XaZNBrJURtCoUT5SUnxFr8s3BzRl+cbzUq8=" crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/js/bootstrap.min.js" integrity="sha384-+sLIOodYLS7CIrQpBjl+C7nPvqq+FbNUBDunl/OZv93DB7Ln/533i8e/mZXLi/P+" crossorigin="anonymous"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/selectize.js/0.15.2/js/selectize.min.js" integrity="sha512-IOebNkvA/HZjMM7MxL0NYeLYEalloZ8ckak+NDtOViP7oiYzG5vn6WVXyrJDiJPhl4yRdmNAG49iuLmhkUdVsQ==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
<script type="text/javascript">
	$("#tecnico").selectize()

	const BtDownload = document.getElementById('BtBaixar');
	const Baixado = document.getElementById('baixado');
	BtDownload.addEventListener('click', function(){
		alert('Arquivo gerado com sucesso, clique em OK para baixar');
		Baixado.click();
	});



</script>
<script type="text/javascript">
function CopyToClipboard(containerid) {
               
	var range = document.createRange();
	range.selectNode(document.getElementById(containerid));
	window.getSelection().removeAllRanges(); // clear current selection
	window.getSelection().addRange(range); // to select text
	document.execCommand("copy");
	window.getSelection().removeAllRanges();// to deselect

	alert("Texto copiado para a área de transferência!")
  
}
</script>
</footer>