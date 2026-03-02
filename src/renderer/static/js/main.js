import { BarChart, createIcons, House, Menu, Monitor, Moon, Settings, Sun, ToyBrick } from "lucide";


const iconsConfig = {
   icons: {
      Menu, House, ToyBrick, BarChart, Settings, Sun, Moon, Monitor
   }
};

// Initial icon creation
document.addEventListener("DOMContentLoaded", () => {
   // TODO: remove
   /** @type {import("@tanstack/table-core").ColumnDef<Record<string, unknown>, any>[]} */
   let columns = [
      {
         header: "Id",
         accessorKey: "id"
      },
      {
         header: "Name",
         accessorKey: "name"
      },
      {
         header: "Email",
         accessorKey: "email"
      },
      {
         header: "Role",
         accessorKey: "role"
      },
      {
         header: "Department",
         accessorKey: "department"
      },
      {
         header: "Status",
         accessorKey: "status"
      },
   ];

   /** @type {Record<string, unknown>[]} */
   let data = [
      {
         id: 26,
         name: "Greg Kihn",
         email: "Myron36@hotmail.com",
         role: "Manager",
         department: "Music",
         status: "Pending",
         children: [
            {
               id: 24,
               name: "Jazmyne Dickinson",
               email: "Westley.Shanahan76@gmail.com",
               role: "Agent",
               department: "Kids",
               status: "Inactive",
               children: []
            }
         ]
      },
      {
         id: 22,
         name: "Gustavo Gislason",
         email: "Shane.Green86@yahoo.com",
         role: "Director",
         department: "Grocery",
         status: "Pending",
         children: [
            {
               id: 28,
               name: "Wendy Hodkiewicz",
               email: "Jadon.Glover54@hotmail.com",
               role: "Analyst",
               department: "Sports",
               status: "Active",
               children: []
            },
            {
               id: 43,
               name: "Sabryna Schinner",
               email: "Eugene_Dibbert62@yahoo.com",
               role: "Analyst",
               department: "Sports",
               status: "Inactive",
               children: []
            }
         ]
      },
      {
         id: 40,
         name: "Eric Reilly",
         email: "Raoul_Graham64@gmail.com",
         role: "Developer",
         department: "Home",
         status: "Active",
         children: []
      },
      {
         id: 0,
         name: "Eloise Gutkowski",
         email: "Rolando.Lesch54@yahoo.com",
         role: "Producer",
         department: "Computers",
         status: "Inactive",
         children: []
      },
      {
         id: 31,
         name: "Kelli Hahn",
         email: "Jovani75@gmail.com",
         role: "Liaison",
         department: "Grocery",
         status: "Pending",
         children: [
            {
               id: 30,
               name: "Selina Littel",
               email: "Kimberly_Wuckert@yahoo.com",
               role: "Supervisor",
               department: "Industrial",
               status: "Pending",
               children: []
            },
            {
               id: 14,
               name: "Kayla Frami I",
               email: "Ardella_Lang@yahoo.com",
               role: "Architect",
               department: "Jewelry",
               status: "Inactive",
               children: []
            },
            {
               id: 50,
               name: "Gertrude Gislason",
               email: "Elias88@yahoo.com",
               role: "Director",
               department: "Music",
               status: "Inactive",
               children: []
            }
         ]
      },
      {
         id: 4,
         name: "Alvah Schimmel",
         email: "Dana31@gmail.com",
         role: "Facilitator",
         department: "Movies",
         status: "Active",
         children: [
            {
               id: 48,
               name: "Willa Lehner",
               email: "Angelica.Bednar14@yahoo.com",
               role: "Manager",
               department: "Sports",
               status: "Pending",
               children: []
            }
         ]
      },
      {
         id: 49,
         name: "Spencer Macejkovic",
         email: "Owen_Sporer@yahoo.com",
         role: "Planner",
         department: "Kids",
         status: "Active",
         children: []
      },
      {
         id: 21,
         name: "Pedro Dicki",
         email: "Sophia34@yahoo.com",
         role: "Representative",
         department: "Shoes",
         status: "Pending",
         children: []
      },
      {
         id: 41,
         name: "Julian Schimmel",
         email: "Iva_Rolfson@yahoo.com",
         role: "Officer",
         department: "Baby",
         status: "Active",
         children: []
      },
      {
         id: 36,
         name: "Alonzo Schroeder I",
         email: "Marjorie82@hotmail.com",
         role: "Executive",
         department: "Automotive",
         status: "Inactive",
         children: [
            {
               id: 17,
               name: "Charlie Tromp III",
               email: "Ashley.Smith10@yahoo.com",
               role: "Agent",
               department: "Industrial",
               status: "Inactive",
               children: []
            },
            {
               id: 3,
               name: "Ross McDermott",
               email: "Herminio29@gmail.com",
               role: "Designer",
               department: "Jewelry",
               status: "Inactive",
               children: []
            },
            {
               id: 35,
               name: "Talon Huel",
               email: "Frances.Ruecker@yahoo.com",
               role: "Officer",
               department: "Automotive",
               status: "Inactive",
               children: []
            }
         ]
      }
   ];

   window.electronAPI.loadPage("home", {
      columns: JSON.stringify(columns),
      data: JSON.stringify(data)
   });
   // TODO: remove end
   // window.electronAPI.loadPage("home")
   createIcons(iconsConfig);
});

// Recreate icons on page change
document.addEventListener("pageChanged", () => {
   createIcons(iconsConfig);
});
