import axios from "axios";
import * as cheerio from "cheerio";
import moment from "moment";

const monthMap: { [date: string]: string } = {
  január: "January",
  február: "February",
  március: "March",
  április: "April",
  május: "May",
  június: "June",
  július: "July",
  augusztus: "August",
  szeptember: "September",
  október: "October",
  november: "November",
  december: "December",
};

export class MNB {
  static async getExchangeRates(_from: Date, _to: Date) {
    // if MNB does not provide rates for the first day (eg. weekend, holiday), we adjust the date range
    const adjustedFrom = new Date(_from.getTime() - 1000 * 60 * 60 * 24 * 7);
    const from = moment(adjustedFrom).format("YYYY.MM.DD.");
    const to = moment(_to).format("YYYY.MM.DD.");

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

    const rates: { date: Date; rate: number }[] = [];

    const $ = cheerio.load(resp.data);
    $("tbody > tr").each((i, el) => {
      const date = $(el).find("td:nth-child(1)").text();
      const rate = $(el).find("td:nth-child(2)").text();

      const dateTranslated = date.replace(
        /január|február|március|április|május|június|július|augusztus|szeptember|október|november|december/g,
        (matched) => monthMap[matched]
      );

      rates.push({
        date: new Date(
          dateTranslated.substring(0, dateTranslated.indexOf(","))
        ),
        rate: parseFloat(rate.replace(/,/g, ".")),
      });
    });

    return rates;
  }
}
