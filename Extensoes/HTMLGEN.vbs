'        @PARANAUERJ DEVELOPEMENT
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







set FSO = CreateObject("Scripting.FileSystemObject")
set Comando = WScript.CreateObject("WScript.Shell")

Class HTMLGEN
    
    private fileName
    private path
    private author
    private description
    private title
    private pageName
    private arq

    Private Sub Class_Initialize(  )
        fileName = "temp.html"
        path = Comando.CurrentDirectory & "\Extensoes\"
        set arq = FSO.CreateTextFile(path & fileName, true)
        arq.Close
    End Sub

    public function initialize(name, ext, aut, des, tit, pname)
        author = aut
        description = des
        title = tit
        pageName = pname

        nfileName = name & "." & ext 
        FSO.deleteFile(path & fileName)

        set arq = FSO.CreateTextFile(path & nfileName, true)

        fileName = nfileName

        arq.WriteLine("<html>")
        arq.WriteLine("<head>")

        arq.WriteLine(" <meta charset=""UTF-8"">")
        arq.WriteLine(" <meta name=""viewport"" content=""width=device-width, initial-scale=1, shrink-to-fit=no"">")
        arq.WriteLine(" <meta name=""description"" content=""" & des & """>")
        arq.WriteLine(" <meta name=""author"" content=""" & aut & """>")
        arq.WriteLine(" <title>" & tit & "</title>")

        arq.WriteLine(" <link href=""../Bibliotecas/Bootstrap/vendor/fontawesome-free/css/all.min.css"" rel=""stylesheet"" type=""text/css"">")
        arq.WriteLine(" <link href=""https://fonts.googleapis.com/css?family=Nunito:200,200i,300,300i,400,400i,600,600i,700,700i,800,800i,900,900i"" rel=""stylesheet"">")
        arq.WriteLine(" <link href=""../Bibliotecas/Bootstrap/css/sb-admin-2.min.css"" rel=""stylesheet"">")
        arq.WriteLine(" <script src=""../Bibliotecas/Bootstrap/vendor/jquery/jquery.min.js""></script>")
        
        arq.WriteLine("</head>")
        arq.WriteLine("<body id=""page-top"">")

        arq.Close

        set arq = FSO.OpenTextFile(path & fileName, 8, true)

    end function

    public function buildBackground()

            arq.WriteLine(" <div id=""content-wrapper"" class=""d-flex flex-column"">")
            arq.WriteLine(" <div id=""content"">")

    end function

    public function buildNavbar(params)

        if params("STATIC") = true then
            arq.WriteLine(" <nav class=""navbar navbar-expand navbar-light bg-white topbar mb-4 static-top shadow"">")
        else
            arq.WriteLine(" <nav class=""navbar navbar-expand navbar-light bg-white topbar mb-4 shadow"">")
        end if

        'arq.WriteLine(" <button id=""sidebarToggleTop"" class=""btn btn-link d-md-none rounded-circle mr-3""><i class=""fa fa-bars""></i></button>")

        if params("SEARCH") = true then

            arq.WriteLine(" <form class=""d-none d-sm-inline-block form-inline mr-auto ml-md-3 my-2 my-md-0 mw-100 navbar-search"">")
            arq.WriteLine("     <div class=""input-group"">")
            arq.WriteLine("         <input type=""text"" class=""form-control bg-light border-0 small"" placeholder=""Search for..."" aria-label=""Search"" aria-describedby=""basic-addon2"">")
            arq.WriteLine("         <div class=""input-group-append"">")
            arq.WriteLine("             <button class=""btn btn-primary"" type=""button"">")
            arq.WriteLine("                 <i class=""fas fa-search fa-sm""></i>")
            arq.WriteLine("             </button>")
            arq.WriteLine("         </div>")
            arq.WriteLine("     </div>")
            arq.WriteLine(" </form>")

        end if

        arq.WriteLine("<ul class=""navbar-nav ml-auto""> <!-- Nav Item - Search Dropdown (Visible Only XS) --><li class=""nav-item dropdown no-arrow d-sm-none""><a class=""nav-link dropdown-toggle"" href=""#"" id=""searchDropdown"" role=""button"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false""><i class=""fas fa-search fa-fw""></i></a><div class=""dropdown-menu dropdown-menu-right p-3 shadow animated--grow-in"" aria-labelledby=""searchDropdown""><form class=""form-inline mr-auto w-100 navbar-search""><div class=""input-group""><input type=""text"" class=""form-control bg-light border-0 small"" placeholder=""Search for..."" aria-label=""Search"" aria-describedby=""basic-addon2""><div class=""input-group-append""><button class=""btn btn-primary"" type=""button""><i class=""fas fa-search fa-sm""></i></button></div></div></form></div></li>")
        
        if params("ICON1") = true then
            buildItemI()
        end if
        if params("ICON2") = true then
            buildItemI()
        end if
        if params("INFO1") = true then
            buildItemII()
        end if
        if params("INFO2") = true then
            buildItemII()
        end if
        
        if params("USER") = true then
            buildItemIII()
        end if

        arq.WriteLine("</ul>")
        arq.WriteLine("</nav>")

    end function

    public function buildTop(btn)

        arq.WriteLine("<div class=""container-fluid"">")

        arq.WriteLine("<div class=""d-sm-flex align-items-center justify-content-between mb-4""><h1 class=""h3 mb-0 text-gray-800"">" & pageName & "</h1>")

        if btn("BUTTON") = true then

            if btn("REDIRECT") = false then
                arq.WriteLine("<a href=""#"" class=""d-none d-sm-inline-block btn btn-sm btn-primary shadow-sm""><i class=""fas fa-download fa-sm text-white-50""></i>")
            else
                arq.WriteLine("<a href=""" & btn("REDIRECT") & """ class=""d-none d-sm-inline-block btn btn-sm btn-primary shadow-sm""><i class=""fas fa-download fa-sm text-white-50""></i>")
            end if

            arq.WriteLine(btn("TEXT"))
            arq.WriteLine("</a>")
            
        end if 
        
        arq.WriteLine("</div>")

    end function

    public function buildFooter()

        arq.WriteLine("</div>")
        arq.WriteLine("<footer class=""sticky-footer bg-white""><div class=""container my-auto""><div class=""copyright text-center my-auto""><span>" & author &"&copy; website <script>var d = new Date();var n = d.getFullYear();document.write(n);</script></span></div></div></footer>")
        arq.WriteLine("</div></div>")
        arq.WriteLine(" <a class=""scroll-to-top rounded"" href=""#page-top""><i class=""fas fa-angle-up""></i></a>")
        arq.WriteLine("<div class=""modal fade"" id=""logoutModal"" tabindex=""-1"" role=""dialog"" aria-labelledby=""exampleModalLabel"" aria-hidden=""true"">    <div class=""modal-dialog"" role=""document"">      <div class=""modal-content"">        <div class=""modal-header"">          <h5 class=""modal-title"" id=""exampleModalLabel"">Ready to Leave?</h5>          <button class=""close"" type=""button"" data-dismiss=""modal"" aria-label=""Close"">            <span aria-hidden=""true"">×</span>          </button>        </div>        <div class=""modal-body"">Select ""Logout"" below if you are ready to end your current session.</div>        <div class=""modal-footer"">          <button class=""btn btn-secondary"" type=""button"" data-dismiss=""modal"">Cancel</button>          <a class=""btn btn-primary"" href=""login.html"">Logout</a>        </div>      </div>    </div>  </div>")

    end function

    public function buildContent(titl)

        arq.WriteLine("<div class=""row"">    <div class=""col-lg-12 mb-4"">        <div class=""card shadow mb-4"">        <div class=""card-header py-3"">          <h6 class=""m-0 font-weight-bold text-primary"">" & titl & "</h6>        </div>        <div class=""card-body"">          <p>...</p>        </div>      </div>    </div>  </div>")

    end function

    public function openContent(titl)

        arq.WriteLine("<div class=""row"">    <div class=""col-lg-12 mb-4"">        <div class=""card shadow mb-4"">        <div class=""card-header py-3"">          <h6 class=""m-0 font-weight-bold text-primary"">" & titl & "</h6>        </div>        <div class=""card-body"">          <p>")

    end function

    public function closeContent()

        arq.WriteLine("</p>        </div>      </div>    </div>  </div>")

    end function

    public function write(str)
        arq.WriteLine(ucFirst(str))

    end function

    public function buildCards(content)

        arq.WriteLine("<div class=""row"">")

        number = UBound(content)





        iteratore = 0

        'msgbox  iteratore & " " & number
        

        while  iteratore <= number
            arq.WriteLine("<div class=""col-md"">    <div class=""card border-left-primary shadow h-100 py-2"">      <div class=""card-body"">        <div class=""row no-gutters align-items-center"">          <div class=""col mr-2"">            <div class=""text-xs font-weight-bold text-primary text-uppercase mb-1"">" & ucFirst(content(iteratore)("TITLE")) & "</div>            <div class=""h5 mb-0 font-weight-bold text-gray-800"">" & ucFirst(content(iteratore)("TEXT")) & "</div>            <small>" & ucFirst(content(iteratore)("DESC")) & "</small>          </div>          <div class=""col-auto"">            <i class=""fas fa-chart-line fa-2x text-gray-300""></i>          </div>        </div>      </div>    </div>  </div>")
             iteratore =  iteratore + 1
        wend

        arq.WriteLine("</div><br>")

    end function

    Private Sub Class_Terminate(  )

        arq.WriteLine("<script src=""../Bibliotecas/Bootstrap/vendor/jquery/jquery.min.js""></script><script src=""../Bibliotecas/Bootstrap/vendor/bootstrap/js/bootstrap.bundle.min.js""></script><script src=""../Bibliotecas/Bootstrap/vendor/jquery-easing/jquery.easing.min.js""></script><script src=""../Bibliotecas/Bootstrap/js/sb-admin-2.min.js""></script>")
        arq.WriteLine("</body>")
        
        arq.WriteLine("</html>")

        arq.Close
        
    End Sub

    public function run(browser)

        browser = UCase(browser)
        if browser = "CHROME" then
            Comando.run browser & " " & " """ & path & filename & """"
        elseif browser = "NATIVE" then
            ' Sem javascript!
           '  conteudo = getText(path & filename)

            Set objExplorer = CreateObject("InternetExplorer.Application")
            With objExplorer
            .Navigate (path & filename)
            .ToolBar = 0
            .StatusBar = 0
            .Left = 100
            .Top = 100
            '.Width = 525
            .Width = 630
            '.Height = 555
            .Height = 800
            .Visible = 1
            .Document.Title = "Pagina WVPP"
            .Document.Body.InnerHTML = conteudo
        End With
        else
            msgbox "Nao ha suporte para o navegador: " & browser
        end if

    end function

    private function buildItemI()

        arq.WriteLine("<!-- Icon -->")
        
        arq.WriteLine("<li class=""nav-item dropdown no-arrow mx-1"">              <a class=""nav-link dropdown-toggle"" href=""#"" id=""alertsDropdown"" role=""button"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">                <i class=""fas fa-bell fa-fw""></i>                <!-- Counter - Alerts -->                <span class=""badge badge-danger badge-counter"">X+</span>              </a>              <!-- Dropdown - Alerts -->              <div class=""dropdown-list dropdown-menu dropdown-menu-right shadow animated--grow-in"" aria-labelledby=""alertsDropdown"">                <h6 class=""dropdown-header"">                  Icon Menu                </h6>                <a class=""dropdown-item d-flex align-items-center"" href=""#"">                  <div class=""mr-3"">                    <div class=""icon-circle bg-primary"">                      <i class=""fas fa-file-alt text-white""></i>                    </div>                  </div>                  <div>                    <div class=""small text-gray-500"">Month day, year</div>                    <span class=""font-weight-bold"">Text1</span>                  </div>                </a>                <a class=""dropdown-item d-flex align-items-center"" href=""#"">                  <div class=""mr-3"">                    <div class=""icon-circle bg-success"">                      <i class=""fas fa-donate text-white""></i>                    </div>                  </div>                  <div>                    <div class=""small text-gray-500"">Month day, year</div>                    Text2                  </div>                </a>                <a class=""dropdown-item d-flex align-items-center"" href=""#"">                  <div class=""mr-3"">                    <div class=""icon-circle bg-warning"">                      <i class=""fas fa-exclamation-triangle text-white""></i>                    </div>                  </div>                  <div>                    <div class=""small text-gray-500"">Month day, year</div>                    Text3                  </div>                </a>                <a class=""dropdown-item text-center small text-gray-500"" href=""#"">Redirect</a>              </div>            </li>")
    
    end function

    private function buildItemII()

        arq.WriteLine("<!-- Info -->")

        arq.WriteLine("<li class=""nav-item dropdown no-arrow mx-1"">    <a class=""nav-link"" aria-expanded=""false"">      <button type=""button"" class=""btn btn-md btn-primary"" disabled>$123.00</button>    </a>  </li>")

    end function

     private function buildItemIII()

        arq.WriteLine("<!-- User -->")
        arq.WriteLine("<div class=""topbar-divider d-none d-sm-block""></div>")
        arq.WriteLine("<li class=""nav-item dropdown no-arrow"">    <a class=""nav-link dropdown-toggle"" href=""#"" id=""userDropdown"" role=""button"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">      <span class=""mr-2 d-none d-lg-inline text-gray-600 small"">Username</span>      <img class=""img-profile rounded-circle"" src=""https://pbs.twimg.com/ext_tw_video_thumb/1221435143480578055/pu/img/cvB0Q7iulnB5x7fV.jpg"">    </a>    <!-- Dropdown - User Information -->    <div class=""dropdown-menu dropdown-menu-right shadow animated--grow-in"" aria-labelledby=""userDropdown"">      <a class=""dropdown-item"" href=""#"">        <i class=""fas fa-user fa-sm fa-fw mr-2 text-gray-400""></i>        Profile      </a>      <a class=""dropdown-item"" href=""#"">        <i class=""fas fa-cogs fa-sm fa-fw mr-2 text-gray-400""></i>        Settings      </a>      <a class=""dropdown-item"" href=""#"">        <i class=""fas fa-list fa-sm fa-fw mr-2 text-gray-400""></i>        Activity Log      </a>      <div class=""dropdown-divider""></div>      <a class=""dropdown-item"" href=""#"" data-toggle=""modal"" data-target=""#logoutModal"">        <i class=""fas fa-sign-out-alt fa-sm fa-fw mr-2 text-gray-400""></i>        Logout      </a>    </div>  </li>")

    end function

End Class