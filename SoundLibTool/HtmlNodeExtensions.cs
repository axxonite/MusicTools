using HtmlAgilityPack;

namespace SoundLibTool
{
	public static class HtmlNodeExtensions
	{
		public static HtmlNodeCollection SelectNodesNoFail(this HtmlNode node, string xpath)
		{
			var result = node.SelectNodes(xpath);
			return result ?? new HtmlNodeCollection(node);
		}

		public static string SelectSingleNodeHtml(this HtmlNode node, string xpath)
		{
			var result = node.SelectSingleNode(xpath);
			return result != null ? result.InnerHtml : "";
		}

		public static string SelectSingleNodeText(this HtmlNode node, string xpath)
		{
			var result = node.SelectSingleNode(xpath);
			return result != null ? result.InnerText : "";
		}
	}
}