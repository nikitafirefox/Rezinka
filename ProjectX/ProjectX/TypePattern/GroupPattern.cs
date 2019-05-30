using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectX.TypePattern
{
    public class GroupPattern
    {
        private int CurrentId { get; set; }
        private int CurrentCount { get; set; }

        private List<PatternParam> PatternParams { get; set; }

        public GroupPattern() {
            PatternParams = new List<PatternParam>();
            Reset();
        }

        public void Reset() {
            CurrentCount = 0;
            CurrentId = 0;
        }

        public void Add(string id, int count) {
            PatternParams.Add(new PatternParam(id, count));
        }

        public string GetIdPatter() {

            if (CurrentCount == PatternParams[CurrentId].Count)
            {
                if (CurrentId < PatternParams.Count - 1)
                {

                    CurrentId++;
                    CurrentCount = 0;
                }
                else
                {
                    CurrentId = 0;
                    CurrentCount = 0;
                }
            }

            CurrentCount++;
            return PatternParams[CurrentId].Id;

        }


    }

    public class PatternParam {

        public string Id { get; set; }
        public int Count { get; set; }

        public PatternParam(string id, int count) {
            Count = count;
            Id = id;
        }

    }

}
