using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace EmailStart.Models
{
    public class HeadersRepository : IHeadersRepository
    {
        private IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["TESTC18"].ConnectionString);

        public int checkMessage(string MessageId)
        {
            DynamicParameters param = new DynamicParameters();
            param.Add("@messageId", MessageId);
           return this.db.Query<int>("sp_oa_CheckMessage", param: param, commandType: CommandType.StoredProcedure).FirstOrDefault();
        }

        public void Mt940txtKaydet(string messageId, string Column)
        {
            DynamicParameters param = new DynamicParameters();
            param.Add("@column", Column);
            param.Add("@messageId", messageId);
            //tekrar proc oluştur.
            db.Query<string>("sp_oa_MT940txt", param: param, commandType: CommandType.StoredProcedure).ToString();

        }
        //list gönderimini kontrol et
        public void UstBilgiKaydet(HeadersModel hm)
        {
            DynamicParameters param = new DynamicParameters();
            param.Add("@From", hm.From);
            param.Add("@Subject", hm.Subject);
            param.Add("@FileName", hm.FileName);
            param.Add("@SentDate", hm.SentDate);
            param.Add("@MessageId", hm.MessageId);

            db.Query<List<HeadersModel>>("sp_oa_ExcelUstBilgiler", param: param, commandType: CommandType.StoredProcedure).ToList();
        }
        public void ZikKaydet(List<BelgeModel> Zartingen, string MessageId)
        {
            int i = 0;
            int z = 1;
            foreach (var item in Zartingen)
            {
                z = 0;
                foreach (var kolon in item.Kolon)
                {
                    z++;
                    if (z == 1)
                    {
                        z = 1;
                        db.Query<string>("insert into tb_oa_exceltablo (Kolon1,MessageId) values(@kolon,@MessageId)", new { kolon, MessageId });
                        i = db.Query<int>("select top 1 Id from tb_oa_exceltablo where MessageId=@MessageId order by id desc", new { MessageId }).SingleOrDefault();

                    }
                    else
                    {
                        string query = "Update tb_oa_exceltablo set Kolon" + z + "=@kolon where Id=@i";
                        db.Query<string>(query, new { kolon, i });
                    }
                }


            }



            // db.Query<string>("insert into tb_cd_ExcelTablo (Kolon1) values(@K)", new { K });
        }
        public void FarkRaporuKaydet(List<BelgeModel> Zartingen, string MessageId)
        {
            int i = 0;
            int z = 1;
            foreach (var item in Zartingen)
            {
                z = 0;
                foreach (var kolon in item.Kolon)
                {
                    z++;
                    if (z == 1)
                    {
                        z = 1;
                        db.Query<string>("insert into tb_oa_exceltablo (Kolon1,MessageId) values(@kolon,@MessageId)", new { kolon, MessageId });
                        i = db.Query<int>("select top 1 Id from tb_oa_exceltablo where MessageId=@MessageId order by id desc", new { MessageId }).SingleOrDefault();

                    }
                    else
                    {
                        string query = "Update tb_oa_exceltablo set Kolon" + z + "=@kolon where Id=@i";
                        db.Query<string>(query, new { kolon, i });
                    }
                }


            }



            // db.Query<string>("insert into tb_cd_ExcelTablo (Kolon1) values(@K)", new { K });
        }

        public int checkStatu(string MessageId)
        {
            DynamicParameters param = new DynamicParameters();
            param.Add("@messageId", MessageId);
            return this.db.Query<int>("sp_oa_CheckStatus", param: param, commandType: CommandType.StoredProcedure).First();
        }

        public void StokKaydet(List<BelgeModel> Zartingen, string MessageId)
        {

            int i = 0;
            int z = 1;
            foreach (var item in Zartingen)
            {
                z = 0;
                foreach (var kolon in item.Kolon)
                {
                    z++;
                    if (z == 1)
                    {
                        z = 1;
                        db.Query<string>("insert into tb_oa_excelstok (Kolon1,MessageId) values(@kolon,@MessageId)", new { kolon, MessageId });
                        i = db.Query<int>("select top 1 Id from tb_oa_exceltablo where MessageId=@MessageId order by id desc", new { MessageId }).SingleOrDefault();

                    }
                    else
                    {
                        string query = "Update tb_oa_exceltablo set Kolon" + z + "=@kolon where Id=@i";
                        db.Query<string>(query, new { kolon, i });
                    }
                }


            }
        }

        public void IadeKaydet(List<BelgeModel> Zartingen, string MessageId)
        {
            int i = 0;
            int z = 1;
            foreach (var item in Zartingen)
            {
                z = 0;
                foreach (var kolon in item.Kolon)
                {
                    z++;
                    if (z == 1)
                    {
                        z = 1;
                        db.Query<string>("insert into tb_oa_excelIade (Kolon1,MessageId) values(@kolon,@MessageId)", new { kolon, MessageId });
                        i = db.Query<int>("select top 1 Id from tb_oa_excelIade where MessageId=@MessageId order by id desc", new { MessageId }).SingleOrDefault();

                    }
                    else
                    {
                        string query = "Update tb_oa_excelIade set Kolon" + z + "=@kolon where Id=@i";
                        db.Query<string>(query, new { kolon, i });
                    }
                }


            }


        }

        public void KpiKaydet(List<BelgeModel> Zartingen, string MessageId)
        {

            int i = 0;
            int z = 1;
            foreach (var item in Zartingen)
            {
                z = 0;
                foreach (var kolon in item.Kolon)
                {
                    z++;
                    if (z == 1)
                    {
                        z = 1;
                        db.Query<string>("insert into tb_oa_exceltablo (Kolon1,MessageId) values(@kolon,@MessageId)", new { kolon, MessageId });
                        i = db.Query<int>("select top 1 Id from tb_oa_exceltablo where MessageId=@MessageId order by id desc", new { MessageId }).SingleOrDefault();

                    }
                    else
                    {
                        string query = "Update tb_oa_exceltablo set Kolon" + z + "=@kolon where Id=@i";
                        db.Query<string>(query, new { kolon, i });
                    }
                }


            }

        }
    }
}