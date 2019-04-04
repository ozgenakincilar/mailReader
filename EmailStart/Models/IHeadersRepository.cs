using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailStart.Models
{
    public interface IHeadersRepository
    {
        void UstBilgiKaydet(HeadersModel hm);
        void Mt940txtKaydet(string messageId, string Column);
        void ZikKaydet(List<BelgeModel> Zartingen, string MessageId);
        int checkMessage(string MessageId);
        int checkStatu(string MessageId);
        void FarkRaporuKaydet(List<BelgeModel> Zartingen, string MessageId);
        void StokKaydet(List<BelgeModel> Zartingen, string MessageId);
        void IadeKaydet(List<BelgeModel> Zartingen, string MessageId);
        void KpiKaydet(List<BelgeModel> Zartingen, string MessageId);
    }
}
