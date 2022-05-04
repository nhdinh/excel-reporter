using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

namespace ExcelReporter
{
    public partial class HorizontalPagination : UserControl, INotifyPropertyChanged
    {
        public static int DEFAULT_PAGE_SIZE = 10;
        private int currentPage = 1;
        private int pageSize = 10;
        private int totalPages = 0;
        private int totalItems = 0;

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public int CurrentPage
        {
            get { return this.currentPage; }
            set
            {
                this.currentPage = value;
                this.NotifyPropertyChanged();
            }
        }

        public int PageSize
        {
            get { return this.pageSize; }
            set
            {
                this.pageSize = value;
                this.totalPages = (int)Math.Ceiling((decimal)this.totalItems / this.pageSize);

                if (this.currentPage > this.totalPages)
                    this.currentPage = this.totalPages;

                this.NotifyPropertyChanged();
            }
        }

        public int TotalItems
        {
            get { return this.totalItems; }
            set
            {
                this.totalItems = value;
                this.totalPages = (int)Math.Ceiling((decimal)this.totalItems / this.pageSize);

                if (this.currentPage > this.totalPages)
                    this.currentPage = this.totalPages;

                this.NotifyPropertyChanged();
            }
        }

        public HorizontalPagination()
        {
            InitializeComponent();
            this.PropertyChanged += HorizontalPagination_PropertyChanged;
        }

        private void HorizontalPagination_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            this.txtShowPage.Text = string.Format("{0}/{1}", this.currentPage, this.totalPages);
        }

        private void HorizontalPagination_Load(object sender, EventArgs e)
        {
            this.txtPageSize.Text = this.pageSize.ToString();
        }

        private void btnNextPage_Click(object sender, EventArgs e)
        {
            if (this.currentPage < this.totalPages)
                this.CurrentPage++;
        }

        private void btnPrevPage_Click(object sender, EventArgs e)
        {
            if (this.currentPage > 1)
                this.CurrentPage--;
        }

        private void btnLastPage_Click(object sender, EventArgs e)
        {
            this.CurrentPage = this.totalPages;
        }

        private void btnFirstPage_Click(object sender, EventArgs e)
        {
            this.CurrentPage = 1;
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            int _pageSize = DEFAULT_PAGE_SIZE;
            try
            {
                _pageSize = int.Parse(this.txtPageSize.Text);
                this.PageSize = _pageSize;
            }
            catch
            {
                this.txtPageSize.Text = this.PageSize.ToString();
            }
        }
    }
}