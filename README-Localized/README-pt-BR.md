# <a name="excel-add-in-typescript-mailmerge"></a>Suplemento Mala Direta TypeScript do Excel

O suplemento Mala Direta TypeScript do Excel conecta-se ao Microsoft Graph, obtém os modelos de email de uma pasta de modelo no Outlook e envia emails de uma lista de destinatários em uma tabela do Excel.

![Página inicial](../readme-images/first_run.PNG)

## <a name="prerequisites"></a>Pré-requisitos

Para executar o exemplo, será necessário:

* Visual Studio 2015
* TypeScript para Microsoft Visual Studio versão 2.0.6.0, no mínimo
* [Node.js](https://nodejs.org/)
* Uma conta de desenvolvedor do Office 365. Caso não tenha uma, [ingresse no Programa para Desenvolvedores do Office 365 e obtenha uma assinatura gratuita de um ano do Office 365](https://aka.ms/devprogramsignup).

## <a name="run-the-add-in"></a>Executar o suplemento

### <a name="register-your-app-in-microsoft-azure"></a>Registrar seu aplicativo no Microsoft Azure

Registre um aplicativo Web no [portal de registro de aplicativos](https://apps.dev.microsoft.com) com a seguinte configuração:

Parâmetro | Valor
---------|--------
Nome | Excel-Add-in-Microsoft-Graph-MailMerge
Tipo | Aplicativo Web e/ou API Web
URL de Logon | https://localhost:44390/index.html
URI da ID do Aplicativo | https://[seu nome de locatário do azure ad].onmicrosoft.com/Excel-Add-in-Microsoft-Graph-MailMerge
URL de Resposta | https://localhost:44390/index.html

Adicione as permissões a seguir:

Aplicativo | Permissões Delegadas
---------|--------
Microsoft Graph | Leitura/Gravação de Email
Microsoft Azure Active Directory | Entrar e ler o perfil do usuário

Salve o aplicativo e anote a *ID do cliente*.

### <a name="set-up-your-environment"></a>Configurar seu ambiente

1. Clone o repositório do GitHub.
3. No Visual Studio, abra o arquivo de solução Excel-Add-in-Microsoft-Graph-MailMerge.sln.

### <a name="update-the-client-id"></a>Atualizar a id do cliente

* Em seu projeto do Visual Studio, abra Excel-Add-in-Microsoft-Graph-MailMergeWeb/src/home/home.ts.
* Atualize '[Insira sua clientId aqui]' com o valor de seu aplicativo do Azure AD.
* Atualize a '[URL de redirecionamento]' com a URL de redirecionamento.

### <a name="run-the-add-in"></a>Executar o suplemento

1. Abra um prompt de comando para o \<diretório de exemplo\>\Excel-Add-in-Microsoft-Graph-MailMergeWeb e execute `npm install`. Quando tiver terminado, execute `npm start`.
2. No Visual Studio, pressione F5 para executar o exemplo.
3. Quando o Excel abrir, selecione o botão de comando **Mala Direta** da guia Página Inicial.

![Botão de comando](../readme-images/command_button.PNG)

4. O painel de tarefas será aberto e você poderá autenticar com as credenciais do Office 365 depois que você clicar em **Entrar com a Microsoft**.
5. Selecione na lista de modelos.

![Selecionar um modelo](../readme-images/select_template.PNG)

6. Examine e edite a lista de destinatários.

![Editar destinatários](../readme-images/mailmerge_table.PNG)

7. Visualize e envie o email.

![Visualizar e enviar emails](../readme-images/preview_send.PNG)

## <a name="questions-and-comments"></a>Perguntas e comentários

Gostaríamos de saber sua opinião sobre este exemplo. Você pode nos enviar suas perguntas e sugestões por meio da seção [Issues](https://github.com/OfficeDev/Excel-Add-in-TypeScript-MailMerge/issues) deste repositório.

As perguntas sobre o desenvolvimento do Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Não deixe de marcar as perguntas ou comentários com [office-addins].

## <a name="additional-resources"></a>Recursos adicionais

* [Exemplos de suplemento do Office](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-add-in)
* [Visão geral da plataforma Suplementos do Office](http://dev.office.com/docs/add-ins/overview/office-add-ins)
* [Introdução aos Suplementos do Office](http://dev.office.com/getting-started/addins)
* [Auxiliares da API JavaScript para Office](https://github.com/OfficeDev/office-js-helpers)

## <a name="copyright"></a>Direitos autorais

Copyright (C) 2016 Microsoft Corporation. Todos os direitos reservados.





