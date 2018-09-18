using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;

namespace equipment_schedule
{
    //предикаты для сортировки массива оборудования
    public class tMachineToolComparerLoad : IComparer
    {
        int IComparer.Compare(Object x, Object y)
        {
            int iCurrLoadX = ((tMachineTool)x).GetCurrentLoad();
            int iCurrLoadY = ((tMachineTool)y).GetCurrentLoad();
            return ((iCurrLoadX < iCurrLoadY) ? (-1) : ((iCurrLoadX > iCurrLoadY) ? (1) : (0)));
        }
    }

    public class tMachineToolComparerID : IComparer
    {
        int IComparer.Compare(Object x, Object y)
        {
            int iIDX = ((tMachineTool)x).GetID();
            int iIDY = ((tMachineTool)y).GetID();
            return ((iIDX < iIDY) ? (-1) : ((iIDX > iIDY) ? (1) : (0)));
        }
    }

    //тип Номенклатура
    public struct tNomenclature
    {
        public int _id;         //ID
        public String _sName;   //Название
    };

    //тип Партия
    public struct tParty
    {
        public int _id;                         //ID
        public tNomenclature _tnNomenclature;   //ссылка на номенклатуру
    };

    //тип Время обработки
    public struct tTime 
    {
        public tNomenclature _tnNomenclature;   //ссылка на номенклатуру
        public int _iTime;                      //время обработки
    };

    //тип Оборудование
    public class tMachineTool
    {
        private int ID;                     //ID
        private String sName;               //Название
        private ArrayList arrTime;          //массив связанных номенклатур с временем обработки (тип элемента tTime)
        public ArrayList arrWorks;          //массив работ (тип tParty)
        private int iCurrentLoad;           //текущий уровень загрузки считая от нулевого момента времени

        //конструктор по умолчанию
        public tMachineTool()               
        {
            ID = -1;
            sName = "";
            arrTime = new ArrayList();
            arrWorks = new ArrayList();
            iCurrentLoad = 0;
        }

        //конструктор с исходными данными, номер и название
        public tMachineTool(int _ID, String _sName)
        {
            this.ID = _ID;
            this.sName = _sName;
            arrTime = new ArrayList();
            arrWorks = new ArrayList();
            iCurrentLoad = 0;
        }

        //возвращает истину, если машина может обрабатывать номенклатуру. на вход задается ID номенклатуры.
        public bool isSuitableForNomenclature(int _id)
        {
            foreach(tTime it in this.arrTime)
            {
                if (it._tnNomenclature._id == _id) return(true);
            }
            return(false);
        }

        //добавляет номенклатуру и время ее обработки в описание оборудования
        //повторное вхождение номенклатуры - отбрасывается
        public int AddTime(tTime _t)
        {
            if (!this.isSuitableForNomenclature(_t._tnNomenclature._id)) arrTime.Add(_t);
            return(0);
        }

        public int GetID()
        {
            return(this.ID);
        }

        public void SetID(int _ID)
        {
            this.ID = _ID;
        }

        public String GetName()
        {
            return(this.sName);
        }

        public void SetName(String _Name)
        {
            this.sName = _Name;
        }

        public int GetCurrentLoad()
        {
            return(this.iCurrentLoad);
        }

        //возвращает время, требуемое на обработку номенклатуры по её ID
        public int GetTimeByNomenclatureID(int _ID)
        {
            foreach(tTime it in this.arrTime)
            {
                if (it._tnNomenclature._id == _ID) return(it._iTime);
            }
            return(-1);
        }

        //записать партию в план обработки машины
        public int AssignJob(tParty Party)
        {
            int iTime = GetTimeByNomenclatureID(Party._tnNomenclature._id);
            if (iTime != -1)
            {
                iCurrentLoad += iTime;
                arrWorks.Add(Party);
                return(0);
            }
            return(-1);
        }

        ~tMachineTool()
        {
        }
    };

    //тип Элемент общего расписания
    public struct tWorkSchedule
    {
        public tParty _Party;       //ссылка на партию
        public tMachineTool _Tool;  //ссылка на машину
        public int _startTime;      //время начала этапа
    };
}
