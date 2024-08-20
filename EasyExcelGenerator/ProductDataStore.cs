namespace EasyExcelGenerator
{
    public class ProductDataStore
    {
        public List<Product> Products { get; private set; }

        public ProductDataStore()
        {
            Products = new List<Product>
            {
                // January
                new Product { Id = 1, Name = "Product A", SalesAmount = 1600m, Month = "January" },
                new Product { Id = 2, Name = "Product B", SalesAmount = 1400m, Month = "January" },
                new Product { Id = 3, Name = "Product C", SalesAmount = 1200m, Month = "January" },
                new Product { Id = 4, Name = "Product D", SalesAmount = 2300m, Month = "January" },

                // February
                new Product { Id = 5, Name = "Product A", SalesAmount = 1300m, Month = "February" },
                new Product { Id = 6, Name = "Product B", SalesAmount = 2400m, Month = "February" },
                new Product { Id = 7, Name = "Product C", SalesAmount = 1600m, Month = "February" },
                new Product { Id = 8, Name = "Product D", SalesAmount = 1900m, Month = "February" },

                // March
                new Product { Id = 9, Name = "Product A", SalesAmount = 1250m, Month = "March" },
                new Product { Id = 10, Name = "Product B", SalesAmount = 2350m, Month = "March" },
                new Product { Id = 11, Name = "Product C", SalesAmount = 1550m, Month = "March" },
                new Product { Id = 12, Name = "Product D", SalesAmount = 1850m, Month = "March" },

                // April
                new Product { Id = 13, Name = "Product A", SalesAmount = 1400m, Month = "April" },
                new Product { Id = 14, Name = "Product B", SalesAmount = 2500m, Month = "April" },
                new Product { Id = 15, Name = "Product C", SalesAmount = 1700m, Month = "April" },
                new Product { Id = 16, Name = "Product D", SalesAmount = 2000m, Month = "April" },

                // May
                new Product { Id = 17, Name = "Product A", SalesAmount = 1500m, Month = "May" },
                new Product { Id = 18, Name = "Product B", SalesAmount = 2600m, Month = "May" },
                new Product { Id = 19, Name = "Product C", SalesAmount = 1800m, Month = "May" },
                new Product { Id = 20, Name = "Product D", SalesAmount = 2100m, Month = "May" },

                // June
                new Product { Id = 21, Name = "Product A", SalesAmount = 1600m, Month = "June" },
                new Product { Id = 22, Name = "Product B", SalesAmount = 2700m, Month = "June" },
                new Product { Id = 23, Name = "Product C", SalesAmount = 1900m, Month = "June" },
                new Product { Id = 24, Name = "Product D", SalesAmount = 2200m, Month = "June" },

                // July
                new Product { Id = 25, Name = "Product A", SalesAmount = 1700m, Month = "July" },
                new Product { Id = 26, Name = "Product B", SalesAmount = 2800m, Month = "July" },
                new Product { Id = 27, Name = "Product C", SalesAmount = 2000m, Month = "July" },
                new Product { Id = 28, Name = "Product D", SalesAmount = 2300m, Month = "July" },

                // August
                new Product { Id = 29, Name = "Product A", SalesAmount = 1800m, Month = "August" },
                new Product { Id = 30, Name = "Product B", SalesAmount = 2900m, Month = "August" },
                new Product { Id = 31, Name = "Product C", SalesAmount = 2100m, Month = "August" },
                new Product { Id = 32, Name = "Product D", SalesAmount = 2400m, Month = "August" },

                // September
                new Product { Id = 33, Name = "Product A", SalesAmount = 1900m, Month = "September" },
                new Product { Id = 34, Name = "Product B", SalesAmount = 3000m, Month = "September" },
                new Product { Id = 35, Name = "Product C", SalesAmount = 2200m, Month = "September" },
                new Product { Id = 36, Name = "Product D", SalesAmount = 2500m, Month = "September" },

                // October
                new Product { Id = 37, Name = "Product A", SalesAmount = 2000m, Month = "October" },
                new Product { Id = 38, Name = "Product B", SalesAmount = 3100m, Month = "October" },
                new Product { Id = 39, Name = "Product C", SalesAmount = 2300m, Month = "October" },
                new Product { Id = 40, Name = "Product D", SalesAmount = 2600m, Month = "October" },

                // November
                new Product { Id = 41, Name = "Product A", SalesAmount = 2100m, Month = "November" },
                new Product { Id = 42, Name = "Product B", SalesAmount = 3200m, Month = "November" },
                new Product { Id = 43, Name = "Product C", SalesAmount = 2400m, Month = "November" },
                new Product { Id = 44, Name = "Product D", SalesAmount = 2700m, Month = "November" },

                // December
                new Product { Id = 45, Name = "Product A", SalesAmount = 2200m, Month = "December" },
                new Product { Id = 46, Name = "Product B", SalesAmount = 3300m, Month = "December" },
                new Product { Id = 47, Name = "Product C", SalesAmount = 2500m, Month = "December" },
                new Product { Id = 48, Name = "Product D", SalesAmount = 2800m, Month = "December" },
            };
        }
    }
}
