using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp3
{
    [Serializable]
    public class Store
    {
        public Store()
        {
            goods = new List<Goods>();
        }
        public List<Goods> goods;
        public Store(List<Goods> goods)
        {
            this.goods = goods;
        }
    }
    [Serializable]
    public class Goods
    {
        public String name;
        public double price;
        public int count;
        public int code;

        public Goods()
        {
            name = "";
            price = 0;
            count = 0;
            code = 0;
        }
        public Goods(String name, double price, int count, int code)
        {
            this.name = name;
            this.price = price;
            this.count = count;
            this.code = code;
        }



    }
}
