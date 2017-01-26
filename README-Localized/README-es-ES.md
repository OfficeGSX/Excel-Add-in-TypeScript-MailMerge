# <a name="excel-add-in-typescript-mailmerge"></a>Complemento Combinar correspondencia de Excel para TypeScript

El complemento Combinar correspondencia de Excel para TypeScript se conecta a Microsoft Graph, obtiene plantillas de correo electrónico de una carpeta de plantillas de Outlook y envía correo desde una lista de destinatarios en una tabla de Excel.

![Página de inicio](../readme-images/first_run.PNG)

## <a name="prerequisites"></a>Requisitos previos

Para ejecutar el ejemplo, necesitará:

* Visual Studio 2015
* TypeScript para Microsoft Visual Studio versión mínima 2.0.6.0
* [Node.js](https://nodejs.org/)
* Una cuenta de desarrollador de Office 365. Si no tiene una, [únase al Programa de desarrolladores de Office 365 y consiga una suscripción gratuita de 1 año a Office 365](https://aka.ms/devprogramsignup).

## <a name="run-the-add-in"></a>Ejecutar el complemento

### <a name="register-your-app-in-microsoft-azure"></a>Registrar la aplicación en Microsoft Azure

Registre una aplicación web en el [Portal de registro de aplicaciones](https://apps.dev.microsoft.com) con la siguiente configuración:

Parámetro | Valor
---------|--------
Nombre | Excel-Add-in-Microsoft-Graph-MailMerge
Tipo | Aplicación web o API web
URL de inicio de sesión | https://localhost:44390/index.html
URI de id. de la aplicación | https://[su nombre de inquilino de Azure AD].onmicrosoft.com/Excel-Add-in-Microsoft-Graph-MailMerge
URL de respuesta | https://localhost:44390/index.html

Agregue los permisos siguientes:

Aplicación | Permisos delegados
---------|--------
Microsoft Graph | Leer o escribir correo
Microsoft Azure Active Directory | Iniciar sesión y leer el perfil del usuario

Guarde la aplicación y anote el *id. de cliente*.

### <a name="set-up-your-environment"></a>Configurar el entorno

1. Clone el repositorio de GitHub.
3. En Visual Studio, abra el archivo de solución Excel-Add-in-Microsoft-Graph-MailMerge.sln.

### <a name="update-the-client-id"></a>Actualizar el id. de cliente

* En su proyecto de Visual Studio, abra Excel-Add-in-Microsoft-Graph-MailMergeWeb/src/home/home.ts.
* Actualice "[Escriba su id. de cliente aquí]" con el valor de su aplicación de Azure AD.
* Actualice la "[URL de redireccionamiento]" con su URL de redireccionamiento.

### <a name="run-the-add-in"></a>Ejecutar el complemento

1. Abra un símbolo del sistema en \<directorio de ejemplo\>\Excel-Add-in-Microsoft-Graph-MailMergeWeb y ejecute `npm install`. Cuando finalice, ejecute `npm start`.
2. En Visual Studio, presione F5 para ejecutar el ejemplo.
3. Cuando Excel se abra, seleccione el botón de comando **Combinar correspondencia** de la pestaña Inicio.

![Botón Comando](../readme-images/command_button.PNG)

4. El panel de tareas se abrirá y podrá autenticarse con las credenciales de Office 365 cuando haga clic en **Iniciar sesión con Microsoft**.
5. Seleccione una plantilla de la lista.

![Seleccionar una plantilla](../readme-images/select_template.PNG)

6. Revise y edite la lista de destinatarios.

![Editar destinatarios](../readme-images/mailmerge_table.PNG)

7. Obtenga una vista previa y envíe el correo electrónico.

![Obtener una vista previa y enviar correos electrónicos](../readme-images/preview_send.PNG)

## <a name="questions-and-comments"></a>Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre este ejemplo. Puede enviarnos sus preguntas y sugerencias a través de la sección [Problemas](https://github.com/OfficeDev/Excel-Add-in-TypeScript-MailMerge/issues) de este repositorio.

Las preguntas generales sobre el desarrollo en Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Asegúrese de que sus preguntas o comentarios se etiquetan con [office-addins].

## <a name="additional-resources"></a>Recursos adicionales

* [Office Add-in samples (Ejemplos de complementos para Office)](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-add-in)
* [Office Add-ins platform overview (Información general sobre la plataforma de complementos para Office)](http://dev.office.com/docs/add-ins/overview/office-add-ins)
* [Get started with Office Add-ins (Introducción a los complementos para Office)](http://dev.office.com/getting-started/addins)
* [Office JavaScript API Helpers (Aplicaciones auxiliares de la API de JavaScript para Office)](https://github.com/OfficeDev/office-js-helpers)

## <a name="copyright"></a>Copyright

Copyright (c) 2016 Microsoft Corporation. Todos los derechos reservados.





