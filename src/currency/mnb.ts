import axios, { AxiosError } from "axios";
import * as cheerio from "cheerio";

export class MNB {
  static async getExchangeRates(year: number) {
    const from = `${year}.01.01.`;
    const to = `${year}.12.31.`;

    var resp = await axios({
      method: "GET",
      url: "https://www.mnb.hu/arfolyam-tablazat",
      params: {
        deviza: "rbCurrencySelect",
        devizaSelected: "USD",
        datefrom: from,
        datetill: to,
        order: 1,
      },
    });

    const rates: { [date: string]: number } = {};

    const $ = cheerio.load(resp.data);
    $("tbody > tr").each((i, el) => {
      const date = $(el).find("td:nth-child(1)").text();
      const rate = $(el).find("td:nth-child(2)").text();

      console.log(date, rate);

      rates[date.substring(0, date.indexOf(","))] = parseFloat(
        rate.replace(/,/g, ".")
      );
    });

    return rates;
  }
}
