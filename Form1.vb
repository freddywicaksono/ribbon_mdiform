Imports DevComponents.DotNetBar
Public Class Form1

    Private Sub btnAddUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddUser.Click
        If (menu_user = False) Then
            BikinMenu(cldUser, TabControl1, "User")
            menu_user = True
        Else
            x = getTabIndex(TabControl1, "User")
            TabControl1.SelectedTabIndex = x
        End If
    End Sub

    Private Sub btnSearchUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchUser.Click
        If (menu_search = False) Then
            BikinMenu(cldSearch, TabControl1, "Search")
            menu_search = True
        Else
            x = getTabIndex(TabControl1, "Search")
            TabControl1.SelectedTabIndex = x
        End If
    End Sub

    Private Sub btnUserReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUserReport.Click
        If (menu_report = False) Then
            BikinMenu(cldReport, TabControl1, "Report")
            menu_report = True
        Else
            x = getTabIndex(TabControl1, "Report")
            TabControl1.SelectedTabIndex = x
        End If
    End Sub

    Private Sub TabControl1_ControlAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.ControlEventArgs) Handles TabControl1.ControlAdded
        TabControl1.SelectedTabIndex = TotalTab - 1
    End Sub

    Private Sub TabControl1_TabItemClose(ByVal sender As Object, ByVal e As DevComponents.DotNetBar.TabStripActionEventArgs) Handles TabControl1.TabItemClose
        Dim itemToRemove As DevComponents.DotNetBar.TabItem = TabControl1.SelectedTab
        If (TotalTab > 2) Then
            TotalTab = TotalTab - 1
        Else
            TotalTab = 0
        End If


        TabControl1.Tabs.Remove(itemToRemove)
        TabControl1.Controls.Remove(itemToRemove.AttachedControl)
        TabControl1.RecalcLayout()


        If (itemToRemove.ToString = "User") Then
            menu_user = False
        ElseIf (itemToRemove.ToString = "Search") Then
            menu_search = False
        ElseIf (itemToRemove.ToString = "Report") Then
            menu_report = False
        Else

        End If
    End Sub

    Private Sub TabControl1_TabItemOpen(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.TabItemOpen
        If (TotalTab = 0) Then
            TotalTab = TotalTab + 2
        Else
            TotalTab = TotalTab + 1
        End If
    End Sub
End Class
