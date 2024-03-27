using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using static System.Net.WebRequestMethods;

namespace BibliotecaSharePoint
{
    internal class Program
    {
        static void Main(string[] args)
        {

            // Configura los detalles de autenticación y las URLs de los sitios de origen y destino
            string sourceSiteUrl = "https://iems.sharepoint.com/sites/ComplyIEMS832";
            string destinationSiteUrl = "https://complyiems.sharepoint.com/sites/UnidaddeInspeccin";
            string usernameorigen = "francisco.rodriguez@iemsamericas.com";
            SecureString passwordorigen = GetSecureString("Comercio11");


            SecureString passworddestino = GetSecureString("Aspnet01");

            // Nombre de la biblioteca que se copiará
            string sourceLibraryName = "Legislacion";
            string destinationLibraryName = "BibliotecaEjemplo";

            // Crea un contexto de cliente para el sitio de origen
            using (var sourceContext = new ClientContext(sourceSiteUrl))
            {
                // Autentica con nombre de usuario y contraseña
                var credentials = new SharePointOnlineCredentials(usernameorigen, passwordorigen);
                sourceContext.Credentials = credentials;

                // Obtén la biblioteca de origen
                var sourceLibrary = sourceContext.Web.Lists.GetByTitle(sourceLibraryName);
                var sourceFiles = sourceLibrary.GetItems(CamlQuery.CreateAllItemsQuery());
                sourceContext.Load(sourceFiles);
                sourceContext.ExecuteQuery();

                // Crea un contexto de cliente para el sitio de destino
                using (var destinationContext = new ClientContext(destinationSiteUrl))
                {
                    // Autentica con nombre de usuario y contraseña
                    var credentialsdestino = new SharePointOnlineCredentials("paco.rdz@complyiems.onmicrosoft.com", passworddestino);
                    sourceContext.Credentials = credentialsdestino;



                    // Crea la biblioteca de destino si no existe
                   var destinationLibrary = destinationContext.Web.Lists.GetByTitle(destinationLibraryName);
                      //  destinationContext.Load(destinationLibrary);
                     //   destinationContext.ExecuteQuery();
    

                    // Copia los archivos de la biblioteca de origen a la biblioteca de destino
                    foreach (var file in sourceFiles)
                    {
                        // Copia el archivo
                        var sourceFile = file.File;
                        sourceContext.Load(sourceFile);
                        sourceContext.ExecuteQuery();

                        var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(sourceContext, sourceFile.ServerRelativeUrl);
                        var fileCreationInfo = new FileCreationInformation();
                        fileCreationInfo.ContentStream = fileInfo.Stream;
                        fileCreationInfo.Url = file.File.Name;
                        var destinationFolder = destinationLibrary.RootFolder;
                        destinationContext.Load(destinationFolder);
                        destinationFolder.Files.Add(fileCreationInfo);
                        destinationContext.ExecuteQuery();
                    }

                    Console.WriteLine("La biblioteca ha sido copiada exitosamente al sitio de destino.");
                }
            }





            /*string urlorigen = "https://iems.sharepoint.com/sites/ComplyIEMS832";
            string librarynameorigen = "Legislacion";

            SecureString passwordorigen = GetSecureString("Comercio11");
            ClientContext clientContextorigen = new ClientContext(urlorigen);
            clientContextorigen.Credentials = new SharePointOnlineCredentials("francisco.rodriguez@iemsamericas.com", passwordorigen);


            string urldestino = "https://complyiems.sharepoint.com/sites/UnidaddeInspeccin";
            string librarynamedestino = "BibliotecaEjemplo";
            SecureString passworddestino = GetSecureString("Aspnet01");
            ClientContext clientContextdestino = new ClientContext(urldestino);
            clientContextdestino.Credentials = new SharePointOnlineCredentials("paco.rdz@complyiems.onmicrosoft.com", passworddestino);

            // Obtener la biblioteca de origen
            List sourceLibrary = clientContextorigen.Web.Lists.GetByTitle(librarynameorigen);
            clientContextorigen.Load(sourceLibrary, l => l.Fields); // Cargar la colección Fields
            clientContextorigen.ExecuteQuery(); // Ejecutar la consulta


            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            ListItemCollection items = sourceLibrary.GetItems(query);

            // Cargar los elementos de la biblioteca de origen
            clientContextorigen.Load(items);
            clientContextorigen.ExecuteQuery();

            // Obtener la biblioteca de destino
            List destinationLibrary = clientContextdestino.Web.Lists.GetByTitle(librarynamedestino);

            // Copiar cada elemento a la biblioteca de destino
            foreach (ListItem item in items)
            {
                // Crear un nuevo elemento en la biblioteca de destino
                ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                ListItem newItem = destinationLibrary.AddItem(newItemInfo);

                // Copiar los valores de los metadatos
                foreach (var field in sourceLibrary.Fields)
                {
                    if (!field.ReadOnlyField && field.InternalName != "Attachments" && field.InternalName != "ContentType")
                    {
                        // Verificar si el valor del campo es válido
                        string fieldValue = item[field.InternalName] != null ? item[field.InternalName].ToString() : string.Empty;
                        if (!string.IsNullOrWhiteSpace(fieldValue))
                        {
                            // Evitar copiar valores con caracteres no válidos
                            if (!ContainsInvalidCharacters(fieldValue))
                            {
                                newItem[field.InternalName] = fieldValue;
                            }
                            else
                            {
                                Console.WriteLine($"El valor del campo {field.InternalName} contiene caracteres no válidos y no se copiará.");
                            }
                        }
                    }

                    // Guardar el nuevo elemento en la biblioteca de destino
                    newItem.Update();
                }



            }*/


        }

            private static bool ContainsInvalidCharacters(string value)
            {
                foreach (char c in value)
                {
                    if (!IsValidXmlChar(c))
                    {
                        return true;
                    }
                }
                return false;
            }

            // Función para verificar si un carácter es válido en XML
            private static bool IsValidXmlChar(char c)
            {
                return c == 0x9 || c == 0xA || c == 0xD ||
                       (c >= 0x20 && c <= 0xD7FF) ||
                       (c >= 0xE000 && c <= 0xFFFD) ||
                       (c >= 0x10000 && c <= 0x10FFFF);
            }












            static void CopiarDatosEntreBibliotecas(ClientContext ccOrigen, ClientContext ccDestino, string nlorigen, string nldestino)
        {
            // Ruta de las carpetas que deseas copiar
            string[] foldersToCopy = { "CONSTITUCION%20MEXICANA", "LEGISLACION%20ESTATAL" };

            // Obtener la biblioteca de origen
            List bibliotecaorigen = ccOrigen.Web.Lists.GetByTitle(nlorigen);
            //CamlQuery query = CamlQuery.CreateAllItemsQuery();
            // ListItemCollection items = bibliotecaorigen.GetItems(query);
            List bibliotecadestino = ccDestino.Web.Lists.GetByTitle(nldestino);
            // Cargar los elementos de la biblioteca de origen
            // ccOrigen.Load(items);
            // ccOrigen.ExecuteQuery();
            // Ejecutar las operaciones de carga en el servidor
            ccOrigen.Load(bibliotecaorigen, l => l.RootFolder);
            ccDestino.Load(bibliotecadestino, l => l.RootFolder);
            ccOrigen.ExecuteQuery();

            ccDestino.ExecuteQuery();

            // Copiar las carpetas especificadas
            foreach (string folderName in foldersToCopy)
            {
                // Obtener la ruta de la carpeta de origen y de destino
                Folder sourceFolder = ccOrigen.Web.GetFolderByServerRelativeUrl(bibliotecaorigen.RootFolder.ServerRelativeUrl + "/" + folderName);
                
                
                Folder destinationFolder = ccDestino.Web.GetFolderByServerRelativeUrl(bibliotecadestino.RootFolder.ServerRelativeUrl + "/" + folderName);

                // Copiar la carpeta y sus archivos
                CopyFolderWithMetadata(sourceFolder, destinationFolder, ccOrigen, ccDestino);
            }

            Console.WriteLine("La copia de las carpetas ha sido completada con éxito.");

            Console.ReadKey();


        }


        static void CopyFolderWithMetadata(Folder sourceFolder, Folder destinationFolder, ClientContext cOrigen, ClientContext cDestino)
        {
            // Cargar la información de la carpeta de origen
            cOrigen.Load(sourceFolder, f => f.Name, f => f.ServerRelativeUrl, f => f.Folders, f => f.Files);
            cOrigen.ExecuteQuery();


            // Crear la carpeta de destino si no existe
            if (!destinationFolder.Exists)
            {
                destinationFolder = destinationFolder.Folders.Add(destinationFolder.Name);
            }
          
            // Copiar los archivos de la carpeta de origen a la carpeta de destino
            foreach (Microsoft.SharePoint.Client.File file in sourceFolder.Files)
            {
                // Cargar los metadatos del archivo de origen
                ListItem sourceItem = file.ListItemAllFields;
                cOrigen.Load(sourceItem);
                cOrigen.ExecuteQuery();

                // Crear un nuevo archivo en la carpeta de destino
                FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(cDestino, file.ServerRelativeUrl);
                Microsoft.SharePoint.Client.File copiedFile = destinationFolder.Files.Add(new FileCreationInformation
                {
                    ContentStream = fileInfo.Stream,
                    Url = file.Name,
                    Overwrite = true
                });
                ListItem destinationItem = copiedFile.ListItemAllFields;

                // Copiar los metadatos del archivo de origen al archivo de destino
                destinationItem["_x00c1_reaTem_x00e1_tica"] = sourceItem["_x00c1_reaTem_x00e1_tica"];
                destinationItem["AutoridadRegulatoria"] = sourceItem["AutoridadRegulatoria"];
                destinationItem["MediodePublicaci_x00f3_n"] = sourceItem["MediodePublicaci_x00f3_n"];
                destinationItem["TipodeDocumento"] = sourceItem["TipodeDocumento"];
                destinationItem["Resumen"] = sourceItem["Resumen"];
                destinationItem["Palabrasclave"] = sourceItem["Palabrasclave"];

                if (sourceItem["Fechadepublicaci_x00f3_n"] == null)
                {
                }
                else
                {
                    DateTime fechaOriginal = ((DateTime)sourceItem["Fechadepublicaci_x00f3_n"]);
                    DateTime nuevaFecha = fechaOriginal.AddDays(1);
                    destinationItem["Fechadepublicaci_x00f3_n"] = nuevaFecha;
                }


                if (sourceItem["_x00da_ltimasReformaPublicada"] == null)
                {
                }
                else
                {
                    DateTime fechaOriginal = ((DateTime)sourceItem["_x00da_ltimasReformaPublicada"]);
                    DateTime nuevaFecha = fechaOriginal.AddDays(1);
                    destinationItem["_x00da_ltimaReformapublicada"] = nuevaFecha;
                }


                destinationItem.Update();
                cDestino.ExecuteQuery();
            }

            // Recursivamente copiar las subcarpetas
            foreach (Folder subFolder in sourceFolder.Folders)
            {
                Folder newSubFolder = destinationFolder.Folders.Add(subFolder.Name);
                CopyFolderWithMetadata(subFolder, newSubFolder,  cOrigen,  cDestino);
            }
        }






        private static SecureString GetSecureString(string password)
        {
            SecureString securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);

            }
            return securePassword;

        }
    }
}
