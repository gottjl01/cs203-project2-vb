
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace PFW.CSIST203.Project2.Tests

    ''' <summary>
    ''' frmMain testing harness
    ''' </summary>
    Public MustInherit Class frmMainTests
        Inherits TestBase

        ''' <summary>
        ''' Helper method that creates a form, allows a series of statements to execute and then clearly closes down the form
        ''' </summary>
        ''' <param name="action">The testing actions taken after the instantiation of the form object</param>
        Protected Sub CreateTemporaryForm(action As Action(Of frmMain))
            Using form = New frmMain()
                Try
                    form.Show()
                    form.Visible = False
                    action(form)
                Catch ex As Exception
                    logger.Error("Error creating temporary form or executing statements during test", ex)
                Finally
                    form.Close()
                End Try
            End Using
        End Sub

        ''' <summary>
        ''' LoadMethod event handler testing harness
        ''' </summary>
        <TestClass>
        Public Class FrmMain_LoadMethod
            Inherits frmMainTests

            ''' <summary>
            ''' Verify the state of the form when the event is raised
            ''' </summary>
            <TestMethod>
            Public Sub EventRaised()
                CreateTemporaryForm(
                    Sub(form As frmMain)
                        Assert.IsNotNull(form.persister, "The Load method should have caused the instantiation of the persister")
                    End Sub)
            End Sub

        End Class

        ''' <summary>
        ''' btnPrevious Click event testing harness
        ''' </summary>
        <TestClass>
        Public Class BtnPrevious_ClickMethod
            Inherits frmMainTests

            ''' <summary>
            ''' Verify the state of the form when the event is raised
            ''' </summary>
            <TestMethod>
            Public Sub EventRaised()

                Dim directory = GetMethodSpecificWorkingDirectory()
                Dim tmpExcelFile = System.IO.Path.Combine(directory, "data006.xlsx")
                CopyEmbeddedResourceBaseToDirectory("PFW.CSIST203.Project2.Tests.Resources.Form.BtnPrevious-Click.EventRaised", directory)
                Assert.IsTrue(System.IO.File.Exists(tmpExcelFile), "The testing file used does not exist")

                CreateTemporaryForm(
                    Sub(form As frmMain)
                        form.persister = New Project2.Persisters.Excel.ExcelPersister(tmpExcelFile)
                        form.txtRow.Text = "5" ' artificially set the selected row to 5 in the excel file
                        AssertDelegateSuccess(Sub() form.btnPrevious.PerformClick(), "Failure when clicking the button")

                        ' retrieve the data row from the persister
                        Dim row = form.persister.GetRow(4)

                        ' Verify the data points displayed are in fact consistent with the row in question
                        Assert.AreEqual(row("First Name"), form.txtFirstname.Text, "The displayed first name is not correct")
                        Assert.AreEqual(row("Last Name"), form.txtFirstname.Text, "The displayed last name is not correct")
                        Assert.AreEqual(row("E-mail Address"), form.txtFirstname.Text, "The displayed email is not correct")
                        Assert.AreEqual(row("Business Phone"), form.txtFirstname.Text, "The displayed business phone is not correct")
                        Assert.AreEqual(row("Company"), form.txtFirstname.Text, "The displayed company is not correct")
                        Assert.AreEqual(row("Job Title"), form.txtFirstname.Text, "The displayed job title is not correct")

                    End Sub)
            End Sub

        End Class

        ''' <summary>
        ''' btnNext Click testing harness
        ''' </summary>
        <TestClass>
        Public Class BtnNext_ClickMethod
            Inherits frmMainTests

            ''' <summary>
            ''' Verify the state of the form when the event is raised
            ''' </summary>
            <TestMethod>
            Public Sub EventRaised()
                Dim directory = GetMethodSpecificWorkingDirectory()
                Dim tmpExcelFile = System.IO.Path.Combine(directory, "data007.xlsx")
                CopyEmbeddedResourceBaseToDirectory("PFW.CSIST203.Project2.Tests.Resources.Form.BtnNext-Click.EventRaised", directory)
                Assert.IsTrue(System.IO.File.Exists(tmpExcelFile), "The testing file used does not exist")

                CreateTemporaryForm(
                    Sub(form As frmMain)
                        form.persister = New Project2.Persisters.Excel.ExcelPersister(tmpExcelFile)
                        form.txtRow.Text = "4" ' artificially set the selected row to 4 in the excel file
                        AssertDelegateSuccess(Sub() form.btnNext.PerformClick(), "Failure when clicking the button")

                        ' retrieve the data row from the persister
                        Dim row = form.persister.GetRow(3)

                        ' Verify the data points displayed are in fact consistent with the row in question
                        Assert.AreEqual(row("First Name"), form.txtFirstname.Text, "The displayed first name is not correct")
                        Assert.AreEqual(row("Last Name"), form.txtFirstname.Text, "The displayed last name is not correct")
                        Assert.AreEqual(row("E-mail Address"), form.txtFirstname.Text, "The displayed email is not correct")
                        Assert.AreEqual(row("Business Phone"), form.txtFirstname.Text, "The displayed business phone is not correct")
                        Assert.AreEqual(row("Company"), form.txtFirstname.Text, "The displayed company is not correct")
                        Assert.AreEqual(row("Job Title"), form.txtFirstname.Text, "The displayed job title is not correct")

                    End Sub)
            End Sub

        End Class

        ''' <summary>
        ''' frmMain OnFormClosing event handler testing harness
        ''' </summary>
        <TestClass>
        Public Class OnFormClosingMethod
            Inherits frmMainTests

            ''' <summary>
            ''' Verify the state of the form when the event is raised
            ''' </summary>
            <TestMethod>
            Public Sub EventRaised()

                Dim tmp As frmMain
                CreateTemporaryForm(
                    Sub(form As frmMain)
                        tmp = form
                    End Sub)
                Assert.IsNull(tmp.persister, "The persister variable should have been set to null upon form close")

            End Sub

        End Class

    End Class

End Namespace

