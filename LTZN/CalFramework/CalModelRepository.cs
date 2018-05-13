using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework
{
    public class CalModelRepository
    {
        public CalModel CreateCalModel(string modelName)
        {
            Guid id = Guid.NewGuid();
            id.ToString();
            CalModel calModel = new CalModel();
            calModel.EID = id.ToString();
            calModel.ModelName = modelName;
            return calModel;
        }

        public List<CalModel> GetALlCalModel()
        {
            return new List<CalModel>();
        }

        public void Remove(CalModel calModel)
        {

        }

        public void Update(CalModel calModel)
        {

        }
    }
}
