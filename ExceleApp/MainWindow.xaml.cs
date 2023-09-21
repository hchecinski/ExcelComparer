using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace ExceleApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Header _header = new Header();
        private List<EntityNumbers> _oldCollection = new List<EntityNumbers>();
        private List<EntityNumbers> _newCollection = new List<EntityNumbers>();
        ScrollViewer sv1, sv2;

        Dictionary<int, string> ColToNumber = new Dictionary<int, string>()
        {
            {0, "A" },
            {1, "B" },
            {2, "C" },
            {3, "D" },
            {4, "E" },
            {5, "F" },
            {6, "G" },
            {7, "H" },
            {8, "I" },
            {9, "J" },
            {10, "K" },
            {11, "L" }
        };


        public MainWindow()
        {
            InitializeComponent();

            SetTriggerForColumnValue();
        }

        private void LoadExcel(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excele (xlsx)|*.xlsx| All files (*.*)|*.*";
            var path = Properties.Settings.Default.defaultPath;
            if(string.IsNullOrEmpty(path))
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            else
            {
                openFileDialog.InitialDirectory = path;
            }

            if(openFileDialog.ShowDialog() == true)
            {

                var fileName = openFileDialog.FileName;
                using(var workbook = new XLWorkbook(fileName))
                {

                    var worksheet = workbook.Worksheets.Worksheet(1);
                    var worksheet2 = workbook.Worksheets.Worksheet(2);

                    _sheet1.Text = worksheet.Name;
                    _sheet2.Text = worksheet2.Name;

                    _header = CreateHeader(worksheet);

                    _oldCollection = CreateRows(worksheet);
                    Binding(_grid, _header, _oldCollection);

                    _newCollection = CreateRows(worksheet2);
                    Binding(_grid2, _header, _newCollection);

                    BrowseChanges(_oldCollection, _newCollection);

                    //Sort(_grid);
                    //Sort(_grid2);

                }
            }
            Properties.Settings.Default.defaultPath = Path.GetDirectoryName(openFileDialog.FileName);
            Properties.Settings.Default.Save();
        }

        private void BrowseChanges(List<EntityNumbers> oldCollection, List<EntityNumbers> newCollection)
        {
            CheckExistFinishedProjects(oldCollection, newCollection);
            CheckExistNewProjects(oldCollection, newCollection);
            CheckChangesInProjects(oldCollection, newCollection);
        }

        private void CheckChangesInProjects(List<EntityNumbers> oldCollection, List<EntityNumbers> newCollection)
        {
            foreach(var newItem in newCollection)
            {
                var oldItem = oldCollection.FirstOrDefault(i => i.A.Value == newItem.A.Value);

                if(oldItem != null)
                {
                    CheckValue(newItem.B, oldItem.B);
                    CheckValue(newItem.C, oldItem.C);
                    CheckValue(newItem.D, oldItem.D);
                    CheckValue(newItem.E, oldItem.E);
                    CheckValue(newItem.F, oldItem.F);
                    CheckValue(newItem.G, oldItem.G);
                    CheckValue(newItem.H, oldItem.H);
                    CheckValue(newItem.I, oldItem.I);
                    CheckValue(newItem.J, oldItem.J);
                    CheckValue(newItem.K, oldItem.K);
                    CheckValue(newItem.L, oldItem.L);

                    if(oldItem.B.IsNewValue 
                        || oldItem.C.IsNewValue
                        || oldItem.D.IsNewValue
                        || oldItem.E.IsNewValue
                        || oldItem.F.IsNewValue
                        || oldItem.G.IsNewValue
                        || oldItem.H.IsNewValue
                        || oldItem.I.IsNewValue
                        || oldItem.J.IsNewValue
                        || oldItem.K.IsNewValue
                        || oldItem.L.IsNewValue)
                    {
                        oldItem.A.IsNewValue = true;
                    }

                }
            }
        }

        private static void CheckValue(Item newItem, Item oldItem)
        {
            if(newItem.Value != oldItem.Value)
            {
                newItem.IsNewValue = true;
            }
        }

        private void CheckExistNewProjects(List<EntityNumbers> oldCollection, List<EntityNumbers> newCollection)
        {
            var newProjects = newCollection.Where(i => !oldCollection.Any(j => j.A.Value == i.A.Value));
            if(newProjects != null && newProjects.Any())
            {
                foreach(var item in newProjects)
                {
                    item.IsNew = true;
                }
            }
        }

        private void CheckExistFinishedProjects(List<EntityNumbers> oldCollection, List<EntityNumbers> newCollection)
        {
            var finishedProjects = oldCollection.Where(i => !newCollection.Any(j => j.A.Value == i.A.Value));

            if(finishedProjects != null && finishedProjects.Any())
            {
                foreach(var item in finishedProjects)
                {
                    item.IsOld = true;
                }
            }
        }

        private void Binding(ListView grid, Header header, List<EntityNumbers> entities)
        {
            grid.ItemsSource = entities;

            var gridView = new GridView();
            foreach(var column in header.Items.OrderBy(i => i.ColumnIndex))
            {
                var letter = ColToNumber[column.ColumnIndex];
                Binding binding = new Binding($"{letter}.Value");

                GridViewColumn gridViewColumn = new GridViewColumn() { Header = column.Value };
                gridViewColumn.Width = letter == "A" ? 200 : 50;

                DataTemplate dataTemplate = new DataTemplate();
                FrameworkElementFactory textBlockFactory = new FrameworkElementFactory(typeof(TextBlock));
                textBlockFactory.SetBinding(TextBlock.TextProperty, binding);
                DataTrigger dataTrigger = new DataTrigger()
                {
                    Binding = new Binding($"{letter}.IsNewValue"),
                    Value = true
                };
                dataTrigger.Setters.Add(new Setter(TextBlock.ForegroundProperty, Brushes.Red));
                dataTrigger.Setters.Add(new Setter(TextBlock.BackgroundProperty, Brushes.Red));
                dataTrigger.Setters.Add(new Setter(TextBlock.FontWeightProperty, FontWeights.Bold));
                dataTrigger.Setters.Add(new Setter(TextBlock.FontSizeProperty, 18.0));

                dataTemplate.VisualTree = textBlockFactory;
                dataTemplate.Triggers.Add(dataTrigger);
                gridViewColumn.CellTemplate = dataTemplate;
                gridView.Columns.Add(gridViewColumn);
            }
            grid.View = gridView;
        }

        private void Sort(ListView grid)
        {
            ICollectionView dataView = CollectionViewSource.GetDefaultView(grid.ItemsSource);

            dataView.SortDescriptions.Clear();
            SortDescription sd = new SortDescription("A.Value", ListSortDirection.Ascending);
            dataView.SortDescriptions.Add(sd);
            dataView.Refresh();
        }

        private List<EntityNumbers> CreateRows(IXLWorksheet worksheet)
        {
            List<EntityNumbers> entities = new List<EntityNumbers>();

            for(int rowIndex = 2; rowIndex < worksheet.RowCount(); rowIndex++)
            {
                EntityNumbers entity = new EntityNumbers();

                entity.A = new Item(worksheet.Row(rowIndex).Cell(1).GetString());
                entity.B = new Item(worksheet.Row(rowIndex).Cell(2).GetString());
                entity.C = new Item(worksheet.Row(rowIndex).Cell(3).GetString());
                entity.D = new Item(worksheet.Row(rowIndex).Cell(4).GetString());
                entity.E = new Item(worksheet.Row(rowIndex).Cell(5).GetString());
                entity.F = new Item(worksheet.Row(rowIndex).Cell(6).GetString());
                entity.G = new Item(worksheet.Row(rowIndex).Cell(7).GetString());
                entity.H = new Item(worksheet.Row(rowIndex).Cell(8).GetString());
                entity.I = new Item(worksheet.Row(rowIndex).Cell(9).GetString());
                entity.J = new Item(worksheet.Row(rowIndex).Cell(10).GetString());
                entity.K = new Item(worksheet.Row(rowIndex).Cell(11).GetString());
                entity.L = new Item(worksheet.Row(rowIndex).Cell(12).GetString());

                if(string.IsNullOrEmpty(entity.A.Value)
                    && string.IsNullOrEmpty(entity.B.Value)
                    && string.IsNullOrEmpty(entity.C.Value))
                {
                    break;
                }

                entities.Add(entity);
            }

            return entities;
        }

        private Header CreateHeader(IXLWorksheet worksheet)
        {
            var header = new Header();

            for(int colIndex = 1; colIndex < worksheet.ColumnCount(); colIndex++)
            {
                string value = worksheet.Row(1).Cell(colIndex).GetString();
                if(string.IsNullOrEmpty(value))
                {
                    if(colIndex == 1)
                    {
                        header.Items.Add(new HeaderItem() { ColumnIndex = colIndex });
                    }
                    else
                    {
                        continue;
                    }
                }
                else
                {
                    header.Items.Add(new HeaderItem() { ColumnIndex = colIndex, Value = value });
                }
            }
            return header;
        }

        private void SetTriggerForColumnValue()
        {
            Style itemContainerStyle = new Style(typeof(ListViewItem));

            DataTrigger triggerOld = new DataTrigger()
            {
                Binding = new Binding("IsOld"),
                Value = true
            };
            triggerOld.Setters.Add(new Setter(ListViewItem.BackgroundProperty, Brushes.LightGray));
            triggerOld.Setters.Add(new Setter(ListViewItem.ForegroundProperty, Brushes.White));
            itemContainerStyle.Triggers.Add(triggerOld);

            DataTrigger triggerNew = new DataTrigger()
            {
                Binding = new Binding("IsNew"),
                Value = true
            };
            triggerNew.Setters.Add(new Setter(ListViewItem.BackgroundProperty, Brushes.LightBlue));
            triggerNew.Setters.Add(new Setter(ListViewItem.ForegroundProperty, Brushes.White));
            itemContainerStyle.Triggers.Add(triggerNew);

            _grid.ItemContainerStyle = itemContainerStyle;
            _grid2.ItemContainerStyle = itemContainerStyle;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            sv1 = VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(this._grid, 0), 0) as ScrollViewer;
            sv2 = VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(this._grid2, 0), 0) as ScrollViewer;

            sv1.ScrollChanged += new ScrollChangedEventHandler(sv1_ScrollChanged);
            sv2.ScrollChanged += new ScrollChangedEventHandler(sv2_ScrollChanged);
        }

        void sv1_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            sv2.ScrollToVerticalOffset(sv1.VerticalOffset);
        }

        void sv2_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            sv1.ScrollToVerticalOffset(sv2.VerticalOffset);
        }
    }
}
