import win32com.client as win32
import datetime
import historico_de_eventos

outlook = win32.Dispatch('Outlook.application')
hora = datetime.datetime.today().strftime('%H:%M')
saudacao = ''
tipo = 'email'

if hora < '12:00':
    saudacao = 'Bom dia'
elif hora >= '12:00' and hora < '18:00' :
    saudacao = 'Boa tarde'      
else:
    saudacao = 'Boa noite'        

def envia_email(nome, anexo, destinatario, copia, assunto, corpo):
    email = outlook.CreateItem(0)
    email.To = destinatario
    email.CC = copia
    email.Subject = assunto
    a1 = '{behavior:url(#default#VML);}'
    a2 = '{font-family:"Cambria Math";ose-1:2 4 5 3 5 4 6 3 2 4;}'
    a3 = '{font-family:Lato;}'
    a4 = '{margin:0cm;t-size:11.0pt;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a5 = '{mso-style-type:personal-compose;t-family:"Calibri",sans-serif;or:windowtext;}'
    a6 = '{mso-style-type:export-only;t-family:"Calibri",sans-serif;-fareast-language:EN-US;}'
    a7 = '{size:612.0pt 792.0pt;gin:70.85pt 3.0cm 70.85pt 3.0cm;}'
    a8 = '{page:WordSection1;}'
    img = r"\\SaoPaulo\Rezende\Jarezende\Administrativo\MIS\Planejamento_MIS\FERRAMENTAS\00-PESSOAL\JOÃO\python\img-assinatura.png"

    email.HTMLBody = f'''
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=iso-8859-1"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><!--[if !mso]><style>v\:* {a1}
    o\:* {a1}
    w\:* {a1}
    .shape {a1}
    </style><![endif]--><style><!--
    /* Font Definitions */
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a2}
    @font-face
    	{a3}
    /* Style Definitions */
    p.MsoNormal, li.MsoNormal, div.MsoNormal
    	{a4}
    span.EstiloDeEmail17
    	{a5}
    .MsoChpDefault
    	{a6}
    @page WordSection1
    	{a7}
    div.WordSection1
    	{a8}
    --></style><!--[if gte mso 9]><xml>
    <o:shapedefaults v:ext="edit" spidmax="1026" />
    </xml><![endif]--><!--[if gte mso 9]><xml>
    <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1" />
    </o:shapelayout></xml><![endif]-->
    </head>
    <body lang=PT-BR link="#0563C1" vlink="#954F72" style='word-wrap:break-word'>
    <div class=WordSection1>
    <p class=MsoNormal>
    <span style='font-size:12.0pt'>
    {nome}, {saudacao}!
    <br>
    <br>Segue em anexo, {corpo}.
    <br>
    <br>Att.</span>
    <o:p>
    </o:p>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <span style='mso-fareast-language:PT-BR'>
    <o:p>&nbsp;</o:p>
    </span>
    </p>
    <p class=MsoNormal style='margin-bottom:8.0pt;line-height:106%'>
    <b>
    <span style='font-size:12.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>João Vitor dos Santos Pinto
    <br>
    </span>
    </b>
    <span style='font-size:10.0pt;line-height:106%;font-family:"Lato",sans-serif;color:#3C3C71;mso-fareast-language:PT-BR'>Assistente de ferramentas
    <br>
    <br>Tel: (11) 3523-0300 | Ramal 2156<br>joao.santos@jarezende.com.br
    <br>www.jarezende.com.br
    <br>
    <br>
    <img src="{img}" width="517" height="117">
    <br>
    <br>
    </span><i>
    <span style='font-size:8.0pt;line-height:106%;font-family:"Calibri Light",sans-serif;color:gray;mso-fareast-language:PT-BR'>
    Esta mensagem pode conter informação confidencial e/ou privilegiada. 
    Se você não for o destinatário ou a pessoa autorizada a receber esta mensagem, não pode usar, 
    copiar ou divulgar as informações nela contidas ou tomar qualquer ação baseada nessas informações. 
    Se você recebeu esta mensagem por engano, por favor, avise imediatamente o remetente, respondendo o e-mail, 
    e em seguida apague-o. 
    Agradecemos sua cooperação.
    <br>
    <br>
    </span>
    </i>
    <i>
    <span style='font-size:8.0pt;line-height:106%;color:gray;mso-fareast-language:PT-BR'>
    This message may contain confidential and/or privileged information. 
    If you are not the address or authorized to receive this for the address, you must not use, copy, 
    disclose or take any action base on this message or any information herein. 
    If you have received this message in error, please advise the sender immediately by reply e-mail and delete this message. 
    Thank you for your cooperation<o:p>
    </o:p></span></i></p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>
    '''

    email.Attachments.Add(anexo)
    email.Send()
    
    historico_de_eventos.importa_historico(anexo, tipo)
    
    #print('e-mail {} enviado'.format(assunto))

