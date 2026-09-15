using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TermStoreTools
{
    // Term Store Tools - full CRUD over SharePoint Online Managed Metadata, using
    // PnP.Framework for interactive auth (replaces the old OfficeDevPnP.Core AuthenticationManager).
    internal class Program
    {
        // Session-level caches so repeated operations against the same term store/group/set/term
        // don't re-prompt or re-fetch from the server every time.
        static TermStore _termStoreCache;
        static readonly Dictionary<string, TermGroup> _groupCache = new Dictionary<string, TermGroup>(StringComparer.OrdinalIgnoreCase);
        static readonly Dictionary<string, TermSet> _termSetCache = new Dictionary<string, TermSet>(StringComparer.OrdinalIgnoreCase);
        static readonly Dictionary<Guid, Term> _termByIdCache = new Dictionary<Guid, Term>();

        static TermStore GetTermStore(ClientContext clientContext)
        {
            if (_termStoreCache != null) return _termStoreCache;

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
            Console.Write("Enter TermStore Name =>");
            TermStore termStore = taxonomySession.TermStores.GetByName(Console.ReadLine());
            clientContext.Load(termStore);
            clientContext.ExecuteQuery();

            _termStoreCache = termStore;
            return termStore;
        }

        static TermGroup GetGroup(ClientContext clientContext, TermStore termStore, string groupName)
        {
            if (_groupCache.TryGetValue(groupName, out TermGroup cached)) return cached;

            TermGroup termGroup = termStore.Groups.GetByName(groupName);
            clientContext.Load(termGroup);
            clientContext.ExecuteQuery();

            _groupCache[groupName] = termGroup;
            return termGroup;
        }

        static TermSet GetTermSet(ClientContext clientContext, TermGroup termGroup, string groupName, string termSetName)
        {
            string key = groupName + "|" + termSetName;
            if (_termSetCache.TryGetValue(key, out TermSet cached)) return cached;

            TermSet termSet = termGroup.TermSets.GetByName(termSetName);
            clientContext.Load(termSet);
            clientContext.ExecuteQuery();

            _termSetCache[key] = termSet;
            return termSet;
        }

        // Prompts for either a Term Id (GUID) - which resolves the Term (and its parent
        // TermSet/Group) directly via TermStore.GetTerm, no need to know the Group/Set names -
        // or falls back to the usual Group -> TermSet -> Term name walk.
        static Term ResolveTerm(ClientContext clientContext, TermStore termStore)
        {
            Console.Write("Enter Term Id (GUID), or leave blank to browse by Group/TermSet/Term name =>");
            string input = Console.ReadLine();

            if (Guid.TryParse(input, out Guid termId))
            {
                if (_termByIdCache.TryGetValue(termId, out Term cachedTerm)) return cachedTerm;

                Term term = termStore.GetTerm(termId);
                clientContext.Load(term,
                    t => t.Name,
                    t => t.Id,
                    t => t.IsDeprecated,
                    t => t.Description,
                    t => t.TermSet.Name,
                    t => t.TermSet.Group.Name);
                clientContext.ExecuteQuery();

                Console.WriteLine("Resolved Term '" + term.Name + "' in TermSet '" + term.TermSet.Name + "' (Group '" + term.TermSet.Group.Name + "')");

                _termByIdCache[termId] = term;
                _groupCache[term.TermSet.Group.Name] = term.TermSet.Group;
                _termSetCache[term.TermSet.Group.Name + "|" + term.TermSet.Name] = term.TermSet;
                return term;
            }

            Console.Write("Enter Group Name =>");
            string groupName = Console.ReadLine();
            TermGroup group = GetGroup(clientContext, termStore, groupName);

            Console.Write("Enter TermSet Name =>");
            string termSetName = Console.ReadLine();
            TermSet termSet = GetTermSet(clientContext, group, groupName, termSetName);

            Console.Write("Enter Term Name =>");
            Term namedTerm = termSet.Terms.GetByName(Console.ReadLine());
            clientContext.Load(namedTerm, t => t.Name, t => t.Id, t => t.IsDeprecated, t => t.Description);
            clientContext.ExecuteQuery();

            _termByIdCache[namedTerm.Id] = namedTerm;
            return namedTerm;
        }

        static void CreateTermGroup(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter New Group Name =>");
            string groupName = Console.ReadLine();

            TermGroup termGroup = termStore.CreateGroup(groupName, Guid.NewGuid());
            clientContext.Load(termGroup);
            clientContext.ExecuteQuery();

            _groupCache[groupName] = termGroup;
            Console.WriteLine("Created Group: " + groupName + " (Id: " + termGroup.Id + ")");
        }

        static void CreateTermSet(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter Group Name =>");
            string groupName = Console.ReadLine();
            TermGroup termGroup = GetGroup(clientContext, termStore, groupName);

            Console.Write("Enter New TermSet Name =>");
            string termSetName = Console.ReadLine();

            TermSet termSet = termGroup.CreateTermSet(termSetName, Guid.NewGuid(), termStore.DefaultLanguage);
            clientContext.Load(termSet);
            clientContext.ExecuteQuery();

            _termSetCache[groupName + "|" + termSetName] = termSet;
            Console.WriteLine("Created TermSet: " + termSetName + " (Id: " + termSet.Id + ")");
        }

        static void CreateTerm(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter Group Name =>");
            string groupName = Console.ReadLine();
            TermGroup termGroup = GetGroup(clientContext, termStore, groupName);

            Console.Write("Enter TermSet Name =>");
            string termSetName = Console.ReadLine();
            TermSet termSet = GetTermSet(clientContext, termGroup, groupName, termSetName);

            Console.Write("Enter Parent Term Id or Name (blank for top-level term) =>");
            string parentTermInput = Console.ReadLine();

            Console.Write("Enter New Term Name =>");
            string newTermName = Console.ReadLine();

            Term newTerm;
            if (string.IsNullOrWhiteSpace(parentTermInput))
            {
                newTerm = termSet.CreateTerm(newTermName, termStore.DefaultLanguage, Guid.NewGuid());
            }
            else
            {
                Term parentTerm = Guid.TryParse(parentTermInput, out Guid parentTermId)
                    ? termStore.GetTerm(parentTermId)
                    : termSet.Terms.GetByName(parentTermInput);
                newTerm = parentTerm.CreateTerm(newTermName, termStore.DefaultLanguage, Guid.NewGuid());
            }
            clientContext.Load(newTerm);
            clientContext.ExecuteQuery();

            _termByIdCache[newTerm.Id] = newTerm;
            Console.WriteLine("Created Term: " + newTermName + " (Id: " + newTerm.Id + ")");
        }

        static void MoveTermSet(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter Source Group Name =>");
            string sourceGroupName = Console.ReadLine();
            TermGroup termGroup = GetGroup(clientContext, termStore, sourceGroupName);

            Console.Write("Enter Source TermSet Name =>");
            string sourceTermSetName = Console.ReadLine();
            TermSet termSet = GetTermSet(clientContext, termGroup, sourceGroupName, sourceTermSetName);

            Console.Write("Enter Destination Group Name =>");
            string destinationGroupName = Console.ReadLine();
            TermGroup destGroup = GetGroup(clientContext, termStore, destinationGroupName);

            termSet.Move(destGroup);
            clientContext.ExecuteQuery();

            _termSetCache.Remove(sourceGroupName + "|" + sourceTermSetName);
            _termSetCache[destinationGroupName + "|" + sourceTermSetName] = termSet;
            Console.WriteLine("Successfully Moved From " + sourceGroupName + " To " + destinationGroupName);
        }

        static void MoveTerm(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Term moveTerm = ResolveTerm(clientContext, termStore);

            Console.Write("Enter Destination Group Name =>");
            string destinationGroupName = Console.ReadLine();
            TermGroup destGroup = GetGroup(clientContext, termStore, destinationGroupName);

            Console.Write("Enter Destination TermSet Name =>");
            string destinationTermSetName = Console.ReadLine();
            TermSet destTermSet = GetTermSet(clientContext, destGroup, destinationGroupName, destinationTermSetName);

            moveTerm.Move(destTermSet);
            clientContext.ExecuteQuery();

            Console.WriteLine("Term moved to " + destinationGroupName + " / " + destinationTermSetName);
        }

        static void RenameTerm(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);
            Term sourceTerm = ResolveTerm(clientContext, termStore);

            Console.Write("Enter New Term Name =>");
            string labelName = Console.ReadLine();
            Console.Write(" New Term LCID 1031(German) 1033(English) =>");
            int lcid = Convert.ToInt32(Console.ReadLine());
            Console.Write("New Term Is Default Yes/ No =>");
            string choice = Console.ReadLine();
            bool isDefaultLabel = choice.ToLower() == "yes" || choice.ToLower() == "y";

            Label changedTerm = sourceTerm.CreateLabel(labelName, lcid, isDefaultLabel);
            clientContext.Load(changedTerm);
            clientContext.ExecuteQuery();
        }

        static void RenameTermSet(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter Group Name =>");
            string groupName = Console.ReadLine();
            TermGroup termGroup = GetGroup(clientContext, termStore, groupName);

            Console.Write("Enter TermSet Name =>");
            string termSetName = Console.ReadLine();
            TermSet termSet = GetTermSet(clientContext, termGroup, groupName, termSetName);

            Console.Write("Enter New TermSet Name =>");
            termSet.Name = Console.ReadLine();
            clientContext.ExecuteQuery();

            Console.WriteLine("TermSet renamed.");
        }

        static void RenameTermGroup(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter Group Name =>");
            string groupName = Console.ReadLine();
            TermGroup termGroup = GetGroup(clientContext, termStore, groupName);

            Console.Write("Enter New Group Name =>");
            termGroup.Name = Console.ReadLine();
            clientContext.ExecuteQuery();

            Console.WriteLine("Group renamed.");
        }

        static void SetTermSetDescription(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter Group Name =>");
            string groupName = Console.ReadLine();
            TermGroup termGroup = GetGroup(clientContext, termStore, groupName);

            Console.Write("Enter TermSet Name =>");
            string termSetName = Console.ReadLine();
            TermSet termSet = GetTermSet(clientContext, termGroup, groupName, termSetName);

            Console.Write("Enter Description =>");
            termSet.Description = Console.ReadLine();
            clientContext.ExecuteQuery();

            Console.WriteLine("TermSet description updated.");
        }

        static void SetTermGroupDescription(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter Group Name =>");
            string groupName = Console.ReadLine();
            TermGroup termGroup = GetGroup(clientContext, termStore, groupName);

            Console.Write("Enter Description =>");
            termGroup.Description = Console.ReadLine();
            clientContext.ExecuteQuery();

            Console.WriteLine("Group description updated.");
        }

        static void SetTermDescription(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);
            Term term = ResolveTerm(clientContext, termStore);

            Console.Write("Enter Description =>");
            string description = Console.ReadLine();

            term.SetDescription(description, termStore.DefaultLanguage);
            clientContext.ExecuteQuery();

            Console.WriteLine("Description updated.");
        }

        static void SetTermAvailability(ClientContext clientContext, bool deprecate)
        {
            TermStore termStore = GetTermStore(clientContext);
            Term term = ResolveTerm(clientContext, termStore);

            term.Deprecate(deprecate);
            clientContext.ExecuteQuery();

            Console.WriteLine(deprecate ? "Term deprecated (hidden from tagging)." : "Term reactivated (available for tagging).");
        }

        static void DeleteTerm(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);
            Term term = ResolveTerm(clientContext, termStore);

            Console.Write("Type 'yes' to confirm permanent deletion of this term =>");
            if (Console.ReadLine().ToLower() != "yes")
            {
                Console.WriteLine("Cancelled.");
                return;
            }

            term.DeleteObject();
            clientContext.ExecuteQuery();

            _termByIdCache.Remove(term.Id);
            Console.WriteLine("Term deleted.");
        }

        static void DeleteTermSet(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter Group Name =>");
            string groupName = Console.ReadLine();
            TermGroup termGroup = GetGroup(clientContext, termStore, groupName);

            Console.Write("Enter TermSet Name to DELETE =>");
            string termSetName = Console.ReadLine();
            TermSet termSet = GetTermSet(clientContext, termGroup, groupName, termSetName);

            Console.Write("Type 'yes' to confirm permanent deletion of this term set and all its terms =>");
            if (Console.ReadLine().ToLower() != "yes")
            {
                Console.WriteLine("Cancelled.");
                return;
            }

            termSet.DeleteObject();
            clientContext.ExecuteQuery();

            _termSetCache.Remove(groupName + "|" + termSetName);
            Console.WriteLine("TermSet deleted.");
        }

        static void DeleteTermGroup(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter Group Name to DELETE =>");
            string groupName = Console.ReadLine();
            TermGroup termGroup = GetGroup(clientContext, termStore, groupName);

            Console.Write("Type 'yes' to confirm permanent deletion of this group and everything in it =>");
            if (Console.ReadLine().ToLower() != "yes")
            {
                Console.WriteLine("Cancelled.");
                return;
            }

            termGroup.DeleteObject();
            clientContext.ExecuteQuery();

            _groupCache.Remove(groupName);
            Console.WriteLine("Group deleted.");
        }

        static void ListTerms(ClientContext clientContext)
        {
            TermStore termStore = GetTermStore(clientContext);

            Console.Write("Enter Group Name =>");
            string groupName = Console.ReadLine();
            TermGroup termGroup = GetGroup(clientContext, termStore, groupName);

            Console.Write("Enter TermSet Name =>");
            string termSetName = Console.ReadLine();
            TermSet termSet = GetTermSet(clientContext, termGroup, groupName, termSetName);

            TermCollection terms = termSet.Terms;
            clientContext.Load(terms, ts => ts.Include(
                t => t.Name,
                t => t.Id,
                t => t.IsDeprecated,
                t => t.Description));
            clientContext.ExecuteQuery();

            foreach (Term term in terms)
            {
                Console.WriteLine(term.Name + "\t" + term.Id + "\t" + (term.IsDeprecated ? "Deprecated" : "Available") + "\t" + term.Description);
            }

            Console.Write("Export to CSV? Enter a file path, or leave blank to skip =>");
            string exportPath = Console.ReadLine();
            if (!string.IsNullOrWhiteSpace(exportPath))
            {
                using (StreamWriter writer = new StreamWriter(exportPath, false))
                {
                    writer.WriteLine("Name,Id,Status,Description");
                    foreach (Term term in terms)
                    {
                        writer.WriteLine("\"" + term.Name.Replace("\"", "\"\"") + "\"," + term.Id + "," +
                            (term.IsDeprecated ? "Deprecated" : "Available") + ",\"" + (term.Description ?? "").Replace("\"", "\"\"") + "\"");
                    }
                }
                Console.WriteLine("Exported to " + exportPath);
            }
        }

        static void Main(string[] args)
        {
            ClientContext clientContext = null;
            try
            {
                Console.Write("Enter Site Url =>");
                string siteUrl = Console.ReadLine();

                using (var authManager = AuthenticationManager.CreateWithInteractiveLogin(AuthenticationManager.CLIENTID_PNPMANAGEMENTSHELL))
                {
                    clientContext = authManager.GetContext(siteUrl);

                    while (true)
                    {
                        Console.WriteLine();
                        Console.WriteLine("Term Store Tools - choose an operation:");
                        Console.WriteLine("Term Group:");
                        Console.WriteLine("  1) Create   2) Rename   3) Update Description   4) Delete");
                        Console.WriteLine("Term Set:");
                        Console.WriteLine("  5) Create   6) Rename   7) Update Description   8) Delete   9) Move to another Group");
                        Console.WriteLine("Term:  (enter a Term Id to skip Group/TermSet lookup)");
                        Console.WriteLine(" 10) Create  11) Rename (add label)  12) Update Description");
                        Console.WriteLine(" 13) Deprecate  14) Reactivate  15) Delete  16) Move to another Term Set");
                        Console.WriteLine("Other:");
                        Console.WriteLine(" 17) List / Export Terms in a Term Set");
                        Console.WriteLine("  q) Quit");
                        Console.Write("=>");
                        string choice = Console.ReadLine();

                        switch (choice)
                        {
                            case "1": CreateTermGroup(clientContext); break;
                            case "2": RenameTermGroup(clientContext); break;
                            case "3": SetTermGroupDescription(clientContext); break;
                            case "4": DeleteTermGroup(clientContext); break;
                            case "5": CreateTermSet(clientContext); break;
                            case "6": RenameTermSet(clientContext); break;
                            case "7": SetTermSetDescription(clientContext); break;
                            case "8": DeleteTermSet(clientContext); break;
                            case "9": MoveTermSet(clientContext); break;
                            case "10": CreateTerm(clientContext); break;
                            case "11": RenameTerm(clientContext); break;
                            case "12": SetTermDescription(clientContext); break;
                            case "13": SetTermAvailability(clientContext, true); break;
                            case "14": SetTermAvailability(clientContext, false); break;
                            case "15": DeleteTerm(clientContext); break;
                            case "16": MoveTerm(clientContext); break;
                            case "17": ListTerms(clientContext); break;
                            default:
                                if (choice != null && choice.ToLower() == "q") return;
                                Console.WriteLine("Unrecognized choice.");
                                break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
            finally
            {
                if (clientContext != null)
                {
                    clientContext.Dispose();
                }
            }
        }
    }
}
