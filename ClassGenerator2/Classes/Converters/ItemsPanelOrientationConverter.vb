Imports System.Globalization
Imports System.Windows.Controls
Imports System.Windows.Data

Namespace WPFConverters

    <ValueConversion(GetType(ItemsPresenter), GetType(Orientation))>
    Public Class ItemsPanelOrientationConverter
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert

            Dim itemsPresenter As ItemsPresenter = TryCast(value, ItemsPresenter)
            If itemsPresenter Is Nothing Then Return Binding.DoNothing
            Dim item As TreeViewItem = TryCast(itemsPresenter.TemplatedParent, TreeViewItem)
            If item Is Nothing Then Return Binding.DoNothing
            Dim isRoot As Boolean = (TypeOf ItemsControl.ItemsControlFromItemContainer(item) Is System.Windows.Controls.TreeView)
            Return If(isRoot, Orientation.Horizontal, Orientation.Vertical)

        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function
    End Class

End Namespace