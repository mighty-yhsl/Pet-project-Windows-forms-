using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace WindowsFormsApp3
{


    public static class Serializator
    {

        public static void SerializeXML<T>(ref T obj, string filename)
        {
            System.Xml.Serialization.XmlSerializer formatter = new System.Xml.Serialization.XmlSerializer(typeof(T));
            using (FileStream fs = new FileStream(filename + ".xml", FileMode.OpenOrCreate))
            {
                formatter.Serialize(fs, obj);

            }
        }

        public static T DeserializeXML<T>(ref T obj, string filename)
        {
            System.Xml.Serialization.XmlSerializer formatter = new System.Xml.Serialization.XmlSerializer(typeof(T));
            using (FileStream fs = new FileStream(filename + ".xml", FileMode.OpenOrCreate))
            {
                obj = (T)formatter.Deserialize(fs);
                return obj;
            }

        }

        public static void SerializeJson<T>(ref T obj, string filename)
        {
            DataContractJsonSerializer writer = new DataContractJsonSerializer(typeof(T));
            using (FileStream fs = new FileStream(filename + ".json", FileMode.OpenOrCreate))
            {
                writer.WriteObject(fs, obj);
            }
        }

        public static T DeserializeJson<T>(ref T obj, string filename)
        {

            DataContractJsonSerializer reader = new DataContractJsonSerializer(typeof(List<Goods>));
            FileStream fs = new FileStream(filename + ".json", FileMode.OpenOrCreate);
            obj = (T)reader.ReadObject(fs);
            fs.Close();
            return obj;
        }

        public static void SerializeBinary<T>(ref T obj, string filename)
        {
            BinaryFormatter writer = new BinaryFormatter();
            using (FileStream fs = new FileStream(filename + ".bin", FileMode.OpenOrCreate))
            {
                writer.Serialize(fs, obj);
            }
        }

        public static T DeserializeBinary<T>(ref T obj, string filename)
        {

            BinaryFormatter reader = new BinaryFormatter();
            FileStream fs = new FileStream(filename + ".bin", FileMode.OpenOrCreate);
            obj = (T)reader.Deserialize(fs);
            fs.Close();
            return obj;
        }

        public static void SerializeConfiguration(Config config)
        {
            System.Xml.Serialization.XmlSerializer formatter = new System.Xml.Serialization.XmlSerializer(typeof(Config));
            using (FileStream fs = new FileStream("config.xml", FileMode.Truncate))
            {
                formatter.Serialize(fs, config);

            }
        }
    }


}
