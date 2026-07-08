import React from "react";

type Service = {
  title: string;
  /** Back image (rotated behind) */
  image: string;
  /** Front image (rotated on top) */
  overlayImage: string;
};

// Data for the service cards.
// Images are Unsplash stock photos; if any fails to load, the onError
// handler swaps in a neutral placeholder so the layout never breaks.
const services: Service[] = [
  {
    title: "Web Development",
    image:
      "https://images.unsplash.com/photo-1461749280684-dccba630e2f6?q=80&w=800&auto=format&fit=crop",
    overlayImage:
      "https://images.unsplash.com/photo-1498050108023-c5249f4df085?q=80&w=800&auto=format&fit=crop",
  },
  {
    title: "Creative Design",
    image:
      "https://images.unsplash.com/photo-1561070791-2526d30994b5?q=80&w=800&auto=format&fit=crop",
    overlayImage:
      "https://images.unsplash.com/photo-1558655146-9f40138edfeb?q=80&w=800&auto=format&fit=crop",
  },
  {
    title: "Branding",
    image:
      "https://images.unsplash.com/photo-1600880292203-757bb62b4baf?q=80&w=800&auto=format&fit=crop",
    overlayImage:
      "https://images.unsplash.com/photo-1611224923853-80b023f02d71?q=80&w=800&auto=format&fit=crop",
  },
  {
    title: "Product Design",
    image:
      "https://images.unsplash.com/photo-1581291518857-4e27b48ff24e?q=80&w=800&auto=format&fit=crop",
    overlayImage:
      "https://images.unsplash.com/photo-1512295767273-ac109ac3acfa?q=80&w=800&auto=format&fit=crop",
  },
];

function handleImageError(
  event: React.SyntheticEvent<HTMLImageElement, Event>,
  fallback: string,
) {
  const target = event.currentTarget;
  target.onerror = null;
  target.src = fallback;
}

// Services Section — "How Can I Help?"
const ServicesSection: React.FC = () => {
  return (
    <div className="bg-white dark:bg-gray-900 min-h-screen w-full flex items-center justify-center font-sans">
      <section className="py-16 sm:py-20 lg:py-24 px-4 sm:px-6 w-full">
        <div className="max-w-5xl mx-auto">
          {/* Section Header */}
          <div className="text-center mb-12 sm:mb-16">
            <h2 className="text-4xl sm:text-5xl lg:text-6xl font-bold text-gray-900 dark:text-white mb-3 sm:mb-4 tracking-tight">
              How Can I Help?
            </h2>
            <p className="text-lg sm:text-xl text-gray-500 dark:text-gray-400 font-light">
              Let&apos;s turn your vision into something amazing.
            </p>
          </div>

          {/* Services Grid */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 lg:gap-8">
            {services.map((service, index) => (
              <div
                key={index}
                className="group bg-gray-50 dark:bg-gray-800/50 rounded-3xl p-6 flex flex-col h-[320px] transition-all duration-300 hover:bg-gray-100 dark:hover:bg-gray-800"
              >
                {/* Image Container */}
                <div className="relative flex-grow flex items-center justify-center mb-4">
                  {/* Back Image */}
                  <img
                    src={service.image}
                    alt={`${service.title} showcase`}
                    loading="lazy"
                    className="absolute w-44 h-auto rounded-lg shadow-md transform -rotate-6 transition-all [transition-duration:400ms] ease-in-out group-hover:rotate-[-10deg] group-hover:scale-105"
                    onError={(e) =>
                      handleImageError(
                        e,
                        "https://placehold.co/512x512/e2e8f0/4a5568?text=Image+1",
                      )
                    }
                  />
                  {/* Front Image */}
                  <img
                    src={service.overlayImage}
                    alt={`${service.title} example`}
                    loading="lazy"
                    className="absolute w-44 h-auto rounded-lg shadow-lg transform rotate-3 transition-all [transition-duration:400ms] ease-in-out group-hover:rotate-[5deg] group-hover:scale-105"
                    onError={(e) =>
                      handleImageError(
                        e,
                        "https://placehold.co/512x512/cbd5e0/2d3748?text=Image+2",
                      )
                    }
                  />
                </div>

                {/* Service Title */}
                <h3 className="text-left text-lg font-medium text-gray-800 dark:text-gray-100 mt-auto">
                  {service.title}
                </h3>
              </div>
            ))}
          </div>
        </div>
      </section>
    </div>
  );
};

export default ServicesSection;
