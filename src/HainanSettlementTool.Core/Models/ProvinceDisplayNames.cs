namespace HainanSettlementTool.Core.Models
{
    public static class ProvinceDisplayNames
    {
        public static string GetName(ProvinceCode province)
        {
            switch (province)
            {
                case ProvinceCode.Chongqing:
                    return "重庆";
                case ProvinceCode.Guangdong:
                    return "广东";
                case ProvinceCode.Hainan:
                    return "海南";
                default:
                    return province.ToString();
            }
        }
    }
}
