'        @PARANAUERJ DEVELOPEMENT
'
'
'	 COMPILADOR VERSÃO BETA 12/07/2018
'
'	   TODOS OS DIREITOS RESERVADOS!
'********************************************







Class Webview

    public width
    public height
    public marginTop
    public marginLeft
    private url
    public nav
    public menuBar
    public toolBar
    public statusBar
    public visible

    Private Sub Class_Initialize(  )
        width = 300
        height = 300
        marginTop = 0
        marginLeft = 0
        menuBar = 0
        toolBar = 0
        statusBar = 0
        visible = 1

        'Trabalha no nav IeObj
        Set nav = CreateObject("InternetExplorer.Application")
        nav.MenuBar = menuBar       
        nav.ToolBar = toolBar
        nav.StatusBar = statusBar
        nav.Left = marginLeft
        nav.Top = marginTop
        nav.Width = width + 30
        nav.Height = height + 30
        nav.visible = visible

    End Sub

    Public function sync()

        nav.MenuBar = menuBar       
        nav.ToolBar = toolBar
        nav.StatusBar = statusBar
        nav.Left = marginLeft
        nav.Top = marginTop
        nav.Width = width + 30
        nav.Height = height + 30
        nav.visible = visible

    end function


    public function run(linke)

        url = linke

        if(isset(url)) then

            nav.Navigate(url)

            Do While nav.Busy Or nav.readyState <> 4
                'Do nothing, wait for the browser to load.
            Loop

            Do While nav.Document.ReadyState <> "complete"
                'Do nothing, wait for the VBScript to load the document of the website.
            Loop

        end if

    end function


    Private Sub Class_Terminate(  )

        nav = nothing

    End Sub

End Class